

const Session = {
  isAuthenticated() {
    return sessionStorage.getItem('qad_authenticated') === 'true';
  },
  getSessionId() {
    return sessionStorage.getItem('qad_session_id') || '';
  },
  getUsername() {
    return sessionStorage.getItem('qad_username') || '';
  },
  getEnvironment() {
    return sessionStorage.getItem('qad_environment') || '';
  },
  clear() {
    sessionStorage.removeItem('qad_authenticated');
    sessionStorage.removeItem('qad_session_id');
    sessionStorage.removeItem('qad_username');
    sessionStorage.removeItem('qad_environment');
  }
};