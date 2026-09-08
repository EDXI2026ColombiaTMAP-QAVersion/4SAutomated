import { next } from '@vercel/functions';

function fromBase64url(value) {
  const normalized = value.replace(/-/g, '+').replace(/_/g, '/');
  return decodeURIComponent(atob(normalized).split('').map(char => `%${(`00${char.charCodeAt(0).toString(16)}`).slice(-2)}`).join(''));
}

async function isValidSession(request) {
  const token = request.headers.get('cookie')?.match(/(?:^|;\s*)teamco_session=([^;]+)/)?.[1];
  if (!token || !process.env.AUTH_SECRET) return false;
  const [payload, providedSignature] = token.split('.');
  if (!payload || !providedSignature) return false;

  const key = await crypto.subtle.importKey(
    'raw', new TextEncoder().encode(process.env.AUTH_SECRET),
    { name: 'HMAC', hash: 'SHA-256' }, false, ['verify']
  );
  const signature = Uint8Array.from(atob(providedSignature.replace(/-/g, '+').replace(/_/g, '/')), char => char.charCodeAt(0));
  const validSignature = await crypto.subtle.verify('HMAC', key, signature, new TextEncoder().encode(payload));
  if (!validSignature) return false;

  try {
    return JSON.parse(fromBase64url(payload)).exp > Date.now();
  } catch {
    return false;
  }
}

export default async function middleware(request) {
  if (await isValidSession(request)) return next();
  const loginUrl = new URL('/login.html', request.url);
  const requestedUrl = new URL(request.url);
  loginUrl.searchParams.set('next', `${requestedUrl.pathname}${requestedUrl.search}`);
  return Response.redirect(loginUrl);
}

export const config = {
  matcher: ['/((?!api/login|api/logout|login\.html|login\.js|styles\.css|favicon\.ico).*)']
};
