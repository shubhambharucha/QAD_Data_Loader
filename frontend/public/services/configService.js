const ConfigService = {
  async load() {
    const res = await fetch(`${API}/api/config`);
    return await res.json();
  },

  async save(payload) {
    const res = await fetch(`${API}/api/save-config`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });

    return await res.json();
  }
};