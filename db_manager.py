import firebase_admin
from firebase_admin import credentials, firestore

if not firebase_admin._apps:
    _cred = None
    try:
        import streamlit as st
        if "firebase" in st.secrets:
            fb = dict(st.secrets["firebase"])
            # TOML may store \n as literal \\n — normalise the private key
            fb["private_key"] = fb["private_key"].replace("\\n", "\n")
            _cred = credentials.Certificate(fb)
    except Exception:
        pass
    if _cred is None:
        _cred = credentials.Certificate("firebase_key.json.json")
    firebase_admin.initialize_app(_cred)

db = firestore.client()
