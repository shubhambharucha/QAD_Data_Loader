(function () {
  if (!Session.isAuthenticated()) {
    window.location.href = '/';
  }
})();