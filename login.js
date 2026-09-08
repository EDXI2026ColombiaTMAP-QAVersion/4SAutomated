const VALID_USERNAME = 'teamco';
const VALID_PASSWORD = '2026teamcowlcteamco2026';
const form = document.querySelector('#loginForm');
const status = document.querySelector('#loginStatus');

form.addEventListener('submit', async event => {
  event.preventDefault();
  const button = form.querySelector('button');
  button.disabled = true;
  status.textContent = '';
  try {
    const formData = new FormData(form);
    const username = formData.get('username');
    const password = formData.get('password');
    if (username !== VALID_USERNAME || password !== VALID_PASSWORD) {
      throw new Error('Invalid username or password.');
    }
    sessionStorage.setItem('teamcoAuthenticated', 'true');
    const next = new URLSearchParams(location.search).get('next') || '/';
    location.assign(next.startsWith('/') ? next : '/');
  } catch (error) {
    status.textContent = error.message;
    button.disabled = false;
  }
});
