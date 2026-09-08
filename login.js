const VALID_USERNAME_HASH = '9118a393e0dba76f8dbb366601b276db8b7d2dba3b0ab5fcfe81d42850145cbe';
const VALID_PASSWORD_HASH = '9ba8d8590adf697bc83e28f459843f798a43983bf8f0d9a8490e16fb506381a7';
const form = document.querySelector('#loginForm');
const status = document.querySelector('#loginStatus');

async function sha256(value) {
  const bytes = new TextEncoder().encode(value);
  const hash = await crypto.subtle.digest('SHA-256', bytes);
  return [...new Uint8Array(hash)].map(byte => byte.toString(16).padStart(2, '0')).join('');
}

form.addEventListener('submit', async event => {
  event.preventDefault();
  const button = form.querySelector('button');
  button.disabled = true;
  status.textContent = '';
  try {
    const formData = new FormData(form);
    const username = formData.get('username');
    const password = formData.get('password');
    const [usernameHash, passwordHash] = await Promise.all([sha256(username), sha256(password)]);
    if (usernameHash !== VALID_USERNAME_HASH || passwordHash !== VALID_PASSWORD_HASH) {
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
