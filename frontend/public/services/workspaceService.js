const WorkspaceService = {
  async getPermissions(sessionId) {
    const res = await fetch(
      `${API}/api/permissions?session_id=${encodeURIComponent(sessionId)}`
    );

    return res;
  },

  async getTemplateEntities() {
    const res = await fetch(`${API}/api/blank-template/entities`);
    return await res.json();
  },

  async logActivity(payload) {
    const res = await fetch(`${API}/api/log-activity`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    return res;
  },

  async validate(payload) {
    return fetch(`${API}/api/validate`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
  },

  async load(payload) {
    return fetch(`${API}/api/load`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
  },

  async downloadBlankTemplate(payload) {
    return fetch(`${API}/api/blank-template`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
  }
};