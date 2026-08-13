const FetchService = {
  async fetchPriceList(payload) {
    return await fetch(`${API}/api/fetch-price-list`, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
  }
};