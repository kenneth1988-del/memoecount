import os
import firebase_admin
from firebase_admin import credentials, firestore


def _init_firebase():
    if firebase_admin._apps:
        return

    # 1. Try Streamlit secrets (Streamlit Cloud and local secrets.toml)
    try:
        import streamlit as st
        if "firebase" in st.secrets:
            fb = dict(st.secrets["firebase"])
            fb["private_key"] = fb["private_key"].replace("\\n", "\n")
            firebase_admin.initialize_app(credentials.Certificate(fb))
            return
    except Exception:
        pass  # st.secrets unavailable — running as a standalone script

    # 2. Fall back to local JSON key file (local development)
    key_file = "firebase_key.json.json"
    if os.path.exists(key_file):
        firebase_admin.initialize_app(credentials.Certificate(key_file))
        return

    # 3. Neither source found — surface a clear error
    try:
        import streamlit as st
        st.error(
            "**Firebase credentials not found.**\n\n"
            "- **Streamlit Cloud:** add a `[firebase]` section to your app secrets.\n"
            "- **Local development:** place `firebase_key.json.json` in the project root."
        )
        st.stop()
    except Exception:
        raise RuntimeError(
            "Firebase credentials not found. "
            "Provide st.secrets [firebase] (cloud) or firebase_key.json.json (local)."
        )


_init_firebase()
db = firestore.client()
