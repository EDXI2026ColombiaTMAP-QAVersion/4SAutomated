const form = document.querySelector('#loginForm');
const status = document.querySelector('#loginStatus');

form.addEventListener('submit', async event => {
  event.preventDefault();
  const button = form.querySelector('button');
  button.disabled = true;
  status.textContent = '';
  try {
    const response = await fetch('/api/login', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(Object.fromEntries(new FormData(form)))
    });
    if (!response.ok) throw new Error('Invalid username or password.');
    const next = new URLSearchParams(location.search).get('next') || '/';
    location.assign(next.startsWith('/') ? next : '/');
  } catch (error) {
    status.textContent = error.message;
    button.disabled = false;
  }
});
