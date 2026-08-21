"""
UNSOS FTS Service Desk Assistant — Streamlit version.

Replaces the earlier client-side EmailJS escalation flow with server-side
email dispatch via the Microsoft Graph API — the approach that actually
works for an Office 365/Exchange Online mailbox like unsos.service@un.org.

WHY NOT SMTP / EmailJS FOR unsos.service@un.org
-------------------------------------------------
Microsoft disabled legacy SMTP AUTH (basic auth) by default for Exchange
Online in 2022, and most UN tenants additionally enforce MFA / Conditional
Access. EmailJS's built-in connectors, and any simple "SMTP + username +
password" approach, depend on exactly that legacy auth path — which is
why the original client-side EmailJS flow failed against a real O365
mailbox. Microsoft Graph's app-only (client credentials) OAuth flow is
the supported modern replacement: no SMTP, no basic auth, works under
MFA/Conditional Access, and is the same mechanism enterprise mail-merge
and ticketing tools use against O365 today.

SETUP REQUIRED (do this with UN ICT, not something end users configure):
1. A UN ICT administrator registers an application in the tenant's Entra
   ID (Azure AD).
2. Grant it Mail.Send **application** permission (not delegated) with
   admin consent.
3. Scope it to the unsos.service@un.org mailbox via an Application Access
   Policy, so the app can only send as that one mailbox, not any mailbox
   in the tenant.
4. Put the resulting tenant_id, client_id, and client_secret into
   Streamlit's secrets (see .streamlit/secrets.toml.example in this repo).

Until that's configured, this app degrades gracefully: it shows a
pre-filled manual email template with a copy button instead of failing
silently.

Author: Gikonyo Ndugu
"""

import time
from datetime import datetime, timezone

import requests
import streamlit as st

try:
    import msal
    MSAL_AVAILABLE = True
except ImportError:
    MSAL_AVAILABLE = False

st.set_page_config(
    page_title="UNSOS FTS Service Desk Assistant",
    page_icon="🛠️",
    layout="centered",
)

# ---------------------------------------------------------------------------
# KNOWLEDGE BASE (ported from the original JS knowledgeBase array)
# ---------------------------------------------------------------------------

def _wifi_steps(location):
    is_somalia = location in ["Mogadishu", "Kismayo", "Baidoa"]
    steps = "1. Check available Wi-Fi networks: look for **UNSOS-CORPORATE** or **UNSOS-GUEST**.\n\n"
    if is_somalia:
        steps += f"**Recommended setup for Somalia field duty stations ({location}):**\n\n"
        steps += "- **Option A (UNSOS-CORPORATE — recommended for staff):** select 'UNSOS-CORPORATE' and log in with your laptop domain credentials — Username: `UNSOS\\your_username`, Password: your laptop password.\n"
        steps += "- **Option B (UNSOS-GUEST):** select 'UNSOS-GUEST' and enter key/password: `UnsosGuestCP@2020`.\n\n"
    else:
        steps += "- **Option A (UNSOS-CORPORATE):** log in using `UNSOS\\username` and your laptop password.\n"
        steps += "- **Option B (UNSOS-GUEST):** connect using password `UnsosGuestCP@2020`.\n\n"
    steps += "2. **Manual configuration (UNSOS-PARTNERS):** Settings > Network & Internet > Wi-Fi > Manage known networks > Forget 'UNSOS-PARTNERS'. Open Network and Sharing Centre > Add new network > WPA2-Enterprise (AES) > uncheck 'Verify server identity'."
    return steps


KNOWLEDGE_BASE = [
    {
        "id": "wifi_config_location_based",
        "category": "NETWORK",
        "triggers": ["wifi", "wi-fi", "wireless", "unsos-partners", "guest", "corporate", "network error", "internet", "connectivity"],
        "get_steps": _wifi_steps,
    },
    {
        "id": "printer_safecom_connection",
        "category": "HARDWARE",
        "triggers": ["printer", "print", "safecom", "paper jam", "toner", "printing"],
        "get_steps": lambda location=None: (
            "1. Verify printer hardware: display active, paper/toner loaded.\n"
            "2. Confirm the network cable is securely plugged in.\n"
            "3. Open Windows **Printers & Scanners**.\n"
            "4. Click **Add device** and select the SafeCom queue (`\\\\printserver\\SafeCom-Print`).\n"
            "5. Tap your UN badge on the reader or enter your SafeCom PIN."
        ),
    },
    {
        "id": "telephony_cisco_phone_troubleshooting",
        "category": "TELEPHONY",
        "triggers": ["cisco", "desk phone", "dial tone", "unregistered", "voip"],
        "get_steps": lambda location=None: (
            "1. Inspect the Ethernet cable from the phone's 'SW' port to the wall jack.\n"
            "2. Verify the PoE display light.\n"
            "3. If status shows 'Unregistered', unplug the cable for 10 seconds to soft-reboot the phone."
        ),
    },
    {
        "id": "telephony_msteams_calling_intermission",
        "category": "TELEPHONY",
        "triggers": ["teams", "ms teams", "calling", "intermission", "mission code", "dial external"],
        "get_steps": lambda location=None: (
            "1. Open MS Teams > **Calls** tab.\n"
            "2. Internal: type staff name or UN extension.\n"
            "3. Intermission: type the Mission Prefix Code + extension.\n"
            "4. If 'Calling disabled' appears, check microphone permissions in Settings > Devices."
        ),
    },
    {
        "id": "fss_umoja_app_access",
        "category": "BUSINESS_APPLICATIONS",
        "triggers": ["fss", "umoja", "portal", "field support", "access denied", "single sign-on"],
        "get_steps": lambda location=None: (
            "1. Open Edge or Chrome in Private/Incognito mode.\n"
            "2. Select **UNSSO / Central Authentication**.\n"
            "3. Log in using your staff email and Unite Identity password.\n"
            "4. Complete MFA on your mobile device."
        ),
    },
    {
        "id": "unite_id_password_reset",
        "category": "USER_ADMINISTRATION",
        "triggers": ["unite id", "password reset", "change password", "forgot password", "unite identity"],
        "get_steps": lambda location=None: (
            "1. Go to uniteid.un.org.\n"
            "2. Enter your Unite ID > Continue.\n"
            "3. Enter current password + new password (12+ chars, uppercase, symbol, number).\n"
            "4. Submit to sync across Active Directory."
        ),
    },
    {
        "id": "mfa_setup_mobile_mail",
        "category": "USER_ADMINISTRATION",
        "triggers": ["mfa", "authenticator", "app password", "mobile mail", "multi-factor"],
        "get_steps": lambda location=None: (
            "1. Log into outlook.office365.com.\n"
            "2. Profile > My Account > Security & privacy.\n"
            "3. 'Additional security verification' > 'Create and manage app passwords'.\n"
            "4. Generate a 16-character app password and use it in your mobile mail app."
        ),
    },
    {
        "id": "outlook_desktop_shared_mailbox",
        "category": "SOFTWARE_AND_EMAIL",
        "triggers": ["shared mailbox", "committee", "additional mailbox", "second email", "open another mailbox"],
        "get_steps": lambda location=None: (
            "1. Desktop: Outlook > File > Add Account > enter target address > Connect.\n"
            "2. Webmail: outlook.office365.com > profile icon > 'Open another mailbox' > enter email.\n"
            "3. If access denied, request delegate access permissions."
        ),
    },
    {
        "id": "windows10_activation",
        "category": "SYSTEM_AND_OS",
        "triggers": ["windows", "activation", "license", "windows 10", "product key"],
        "get_steps": lambda location=None: (
            "1. Start Menu > Settings > Update & Security > Activation.\n"
            "2. Click 'Change product key'.\n"
            "3. Enter the official UNSOS volume license key > Activate."
        ),
    },
]

GENERIC_FIRST_PHASE_TIPS = (
    "1. Restart the affected application or device.\n"
    "2. Confirm you're connected to UNSOS-CORPORATE or UNSOS-GUEST Wi-Fi.\n"
    "3. Confirm you're signed in with your UN email and current Unite ID password.\n"
    "4. Note the exact error message on screen — it helps the technician diagnose faster."
)


def match_knowledge_base(text):
    lower = text.lower()
    for item in KNOWLEDGE_BASE:
        if any(trig in lower for trig in item["triggers"]):
            return item
    return None


# ---------------------------------------------------------------------------
# EMAIL DISPATCH — Microsoft Graph API
# ---------------------------------------------------------------------------

def send_escalation_graph(ticket):
    """Send the escalation email via Microsoft Graph (app-only OAuth).
    Returns (success: bool, error_message: str | None)."""
    if not MSAL_AVAILABLE:
        return False, "The 'msal' package is not installed."

    try:
        tenant_id = st.secrets["graph"]["tenant_id"]
        client_id = st.secrets["graph"]["client_id"]
        client_secret = st.secrets["graph"]["client_secret"]
        sender_mailbox = st.secrets["graph"].get("sender_mailbox", "unsos.service@un.org")
    except Exception:
        return False, "Graph API credentials are not configured in st.secrets (see .streamlit/secrets.toml.example)."

    try:
        app = msal.ConfidentialClientApplication(
            client_id,
            authority=f"https://login.microsoftonline.com/{tenant_id}",
            client_credential=client_secret,
        )
        token_result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" not in token_result:
            return False, f"Token acquisition failed: {token_result.get('error_description', 'unknown error')}"

        access_token = token_result["access_token"]

        body_text = (
            f"Ticket ID: {ticket['ticket_id']}\n"
            f"Requester: {ticket['full_name']} ({ticket['email']})\n"
            f"Duty Station: {ticket['location']} - {ticket['building_block']}\n"
            f"Category: {ticket['category']}\n"
            f"Issue: {ticket['query']}\n"
            f"Escalation reason: {ticket['reason']}\n"
            f"Routing group: {ticket['routing_group']}\n"
            f"Submitted: {ticket['timestamp']}\n"
        )

        message = {
            "message": {
                "subject": f"[Escalation Ticket] #{ticket['ticket_id']} - IT Support Request",
                "body": {"contentType": "Text", "content": body_text},
                "toRecipients": [{"emailAddress": {"address": "unsos.service@un.org"}}],
                "replyTo": [{"emailAddress": {"address": ticket["email"]}}],
            },
            "saveToSentItems": "true",
        }

        resp = requests.post(
            f"https://graph.microsoft.com/v1.0/users/{sender_mailbox}/sendMail",
            headers={"Authorization": f"Bearer {access_token}", "Content-Type": "application/json"},
            json=message,
            timeout=15,
        )
        if resp.status_code == 202:
            return True, None
        return False, f"Graph API returned {resp.status_code}: {resp.text[:300]}"
    except Exception as e:
        return False, str(e)


def build_manual_email_template(ticket):
    return (
        f"Subject: [Escalation Ticket] #{ticket['ticket_id']} - IT Support Request\n\n"
        f"Dear UNSOS Service Desk,\n\n"
        f"Please note that my automated service desk escalation failed to transmit directly, "
        f"so I am submitting this request manually.\n\n"
        f"Ticket ID: #{ticket['ticket_id']}\n"
        f"Requester Name: {ticket['full_name']}\n"
        f"Requester Email: {ticket['email']}\n"
        f"Duty Station: {ticket['location']} - {ticket['building_block']}\n"
        f"Category: {ticket['category']}\n"
        f"Status: Pending Technician Assignment\n\n"
        f"Issue Description:\n{ticket['query']}\n\n"
        f"Kindly assist in routing this to the appropriate technical support group.\n\n"
        f"Best regards,\n{ticket['full_name']}"
    )


def make_ticket(user, category, query, reason):
    is_somalia = user["location"] in ["Mogadishu", "Kismayo", "Baidoa"]
    routing_group = (
        f"Somalia Field IT Support ({user['location']})"
        if is_somalia
        else "UNSOS Central Service Desk (Nairobi)"
    )
    return {
        "ticket_id": "UNSOS-" + str(int(time.time()))[-6:],
        "full_name": user["full_name"],
        "email": user["email"],
        "location": user["location"],
        "building_block": user["building_block"],
        "category": category,
        "query": query or "No specific query logged — user escalated directly.",
        "reason": reason,
        "routing_group": routing_group,
        "timestamp": datetime.now(timezone.utc).isoformat(),
    }


# ---------------------------------------------------------------------------
# SESSION STATE
# ---------------------------------------------------------------------------

defaults = {
    "registered": False,
    "session_user": {},
    "chat_history": [],
    "last_query": "",
    "last_category": "GENERAL_IT",
    "pending_manual_ticket": None,
}
for key, val in defaults.items():
    if key not in st.session_state:
        st.session_state[key] = val

st.title("🛠️ UNSOS FTS Service Desk Assistant")
st.caption("Field Support & Ticket Classifier — Streamlit edition")

if not MSAL_AVAILABLE:
    st.info(
        "Note: the `msal` package isn't installed in this environment, so automated escalation "
        "will fall back to the manual email template below. Run `pip install -r requirements.txt` "
        "to enable it.",
        icon="ℹ️",
    )

# ---------------------------------------------------------------------------
# STEP 1 — Registration form
# ---------------------------------------------------------------------------
if not st.session_state.registered:
    st.subheader("Please provide your details to start support")
    with st.form("registration_form"):
        full_name = st.text_input("Full name")
        email = st.text_input("UN email address")
        location = st.selectbox("Duty Station / Location", ["", "Nairobi", "Mogadishu", "Kismayo", "Baidoa"])
        building_block = st.text_input("Building / Block / Office")
        submitted = st.form_submit_button("Start Support Chat")

        if submitted:
            if not (full_name and email and location and building_block):
                st.error("Please complete all required fields before starting the chat.")
            else:
                st.session_state.session_user = {
                    "full_name": full_name,
                    "email": email,
                    "location": location,
                    "building_block": building_block,
                }
                st.session_state.registered = True
                st.session_state.chat_history.append(
                    {
                        "role": "bot",
                        "text": (
                            f"Hello {full_name}!\n\nWelcome to UNSOS FTS Service Desk Assistant.\n\n"
                            f"Location registered: **{location} ({building_block})**.\n\nHow can I help you today?"
                        ),
                    }
                )
                st.rerun()

# ---------------------------------------------------------------------------
# STEP 2 — Chat interface
# ---------------------------------------------------------------------------
else:
    user = st.session_state.session_user

    for msg in st.session_state.chat_history:
        with st.chat_message("assistant" if msg["role"] == "bot" else "user"):
            st.markdown(msg["text"])

    # Manual fallback template, shown when the last automated send failed
    if st.session_state.pending_manual_ticket:
        ticket = st.session_state.pending_manual_ticket
        with st.chat_message("assistant"):
            st.warning(
                f"**[TICKET NOT CONFIRMED - #{ticket['ticket_id']}]**\n\n"
                f"We couldn't confirm the notification to {ticket['routing_group']} went through."
            )
            template_text = build_manual_email_template(ticket)
            st.text_area(
                "Manual Email Escalation Template — copy and send to UNSOS.service@un.org:",
                template_text,
                height=280,
                key=f"manual_{ticket['ticket_id']}",
            )
            if st.button("Retry automated send", key=f"retry_{ticket['ticket_id']}"):
                with st.spinner("Retrying..."):
                    success, error = send_escalation_graph(ticket)
                if success:
                    st.session_state.chat_history.append(
                        {
                            "role": "bot",
                            "text": (
                                f"**[TICKET CREATED - #{ticket['ticket_id']}]**\n\n"
                                f"Your request has been escalated to {ticket['routing_group']}.\n\n"
                                f"Status: Confirmed — notification sent."
                            ),
                        }
                    )
                    st.session_state.pending_manual_ticket = None
                    st.rerun()
                else:
                    st.error(f"Retry failed: {error}")

    query = st.chat_input("Type your issue (e.g., Wi-Fi, printer, MFA, Teams, FSS)...")
    if query:
        st.session_state.chat_history.append({"role": "user", "text": query})
        st.session_state.last_query = query
        matched = match_knowledge_base(query)

        if matched:
            st.session_state.last_category = matched["category"]
            steps = matched["get_steps"](user["location"])
            bot_text = (
                f"**[Category: {matched['category']}]**\n\n1st-Phase Resolution Steps:\n\n{steps}\n\n"
                f"Still need help? Use the **Escalate to IT Support** button below."
            )
        else:
            st.session_state.last_category = "UNCLASSIFIED"
            bot_text = (
                f"**[Category: UNCLASSIFIED]**\n\nI couldn't find an exact match for that in the "
                f"knowledge base. A few general steps to try first:\n\n{GENERIC_FIRST_PHASE_TIPS}\n\n"
                f"Still need help? Use the **Escalate to IT Support** button below."
            )

        st.session_state.chat_history.append({"role": "bot", "text": bot_text})
        st.rerun()

    st.divider()
    col1, col2 = st.columns([3, 1])
    with col1:
        if st.button("🚨 Escalate to IT Support", type="primary", use_container_width=True):
            ticket = make_ticket(
                user,
                st.session_state.last_category,
                st.session_state.last_query,
                "User requested manual escalation.",
            )
            with st.spinner(f"Creating ticket #{ticket['ticket_id']} and notifying {ticket['routing_group']}..."):
                success, error = send_escalation_graph(ticket)

            if success:
                st.session_state.chat_history.append(
                    {
                        "role": "bot",
                        "text": (
                            f"**[TICKET CREATED - #{ticket['ticket_id']}]**\n\n"
                            f"Your request has been escalated to {ticket['routing_group']}.\n\n"
                            f"- Requester: {user['full_name']} ({user['email']})\n"
                            f"- Location: {user['location']} - {user['building_block']}\n"
                            f"- Status: Confirmed — notification sent.\n\n"
                            f"Please keep ticket #{ticket['ticket_id']} for reference."
                        ),
                    }
                )
                st.session_state.pending_manual_ticket = None
            else:
                st.session_state.chat_history.append(
                    {
                        "role": "bot",
                        "text": (
                            f"**[TICKET NOT CONFIRMED - #{ticket['ticket_id']}]**\n\n"
                            f"Automated dispatch failed ({error}). See the manual template below."
                        ),
                    }
                )
                st.session_state.pending_manual_ticket = ticket
            st.rerun()

    with col2:
        if st.button("Start over", use_container_width=True):
            for key, val in defaults.items():
                st.session_state[key] = val
            st.rerun()
