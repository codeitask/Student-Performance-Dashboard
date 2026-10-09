import streamlit as st

def check_login():
    """Loops through all users in secrets to validate credentials."""
    try:
        users = st.secrets["users"]
        for user in users:
            if st.session_state.username == user["username"] and st.session_state.password == user["password"]:
                st.session_state.logged_in = True
                st.session_state.user_role = user["role"]
                st.session_state.current_user = user["username"]
                del st.session_state.username
                del st.session_state.password
                return
        st.error("❌ Incorrect username or password")
    except KeyError:
        st.error("🔒 App Setup Incomplete: 'users' array is missing in Secrets.")

def logout():
    """Logs out the user and clears all session states."""
    st.session_state.logged_in = False
    st.session_state.user_role = None
    st.session_state.current_user = None

def show_login_page():
    """Renders the login screen. Returns True if logged in, False otherwise."""
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False

    if not st.session_state.logged_in:
        st.title("🔒 Corporate Dashboard Login")
        st.text_input("Username", key="username")
        st.text_input("Password", type="password", key="password")
        st.button("Log In", on_click=check_login, type="primary")
        return False
        
    # User is logged in, show a personalized sidebar
    st.sidebar.write(f"👤 Logged in as: **{st.session_state.current_user}** ({st.session_state.user_role})")
    st.sidebar.button("Log Out", on_click=logout)
    return True
