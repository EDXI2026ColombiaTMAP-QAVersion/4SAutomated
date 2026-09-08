const crypto = require('crypto');

function base64url(value) {
  return Buffer.from(value).toString('base64url');
}

function sign(payload) {
  return crypto.createHmac('sha256', process.env.AUTH_SECRET)
    .update(payload)
    .digest('base64url');
}

module.exports = (req, res) => {
  if (req.method !== 'POST') return res.status(405).json({ error: 'Method not allowed' });

  const { username, password } = req.body || {};
  if (username !== process.env.TEAMCO_USER || password !== process.env.TEAMCO_PASSWORD) {
    return res.status(401).json({ error: 'Invalid credentials' });
  }

  const expires = Date.now() + 8 * 60 * 60 * 1000;
  const payload = base64url(JSON.stringify({ user: username, exp: expires }));
  const token = `${payload}.${sign(payload)}`;
  res.setHeader('Set-Cookie', `teamco_session=${token}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=28800`);
  return res.status(200).json({ ok: true });
};
