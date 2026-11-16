const productionHost = 'dentalcirclegroup.com';
const productionOrigin = `https://${productionHost}`;
const previewOrigin = 'https://preview.dentalcirclegroup.com';

const publicPages = new Set([
  '/',
  '/about',
  '/appointment',
  '/blog',
  '/contact',
  '/faq',
  '/gallery',
  '/services',
  '/testimonials',
  '/services/cosmetic-dentistry',
  '/services/dental-implants',
  '/services/emergency-dentistry',
  '/services/family-dentistry',
  '/services/invisalign',
  '/services/oral-surgery',
  '/services/orthodontics',
  '/services/periodontics',
  '/services/preventative-dentistry',
  '/services/prosthodontics',
  '/services/root-canal-treatment',
  '/services/smile-makeover',
  '/services/teeth-whitening',
  '/services/veneers',
]);

export default {
  async fetch(request) {
    const url = new URL(request.url);

    if (url.hostname !== productionHost) {
      return new Response('Not found', { status: 404 });
    }

    const upstream = await fetch(request);
    const contentType = upstream.headers.get('Content-Type') ?? '';

    if (
      request.method !== 'GET' ||
      upstream.status !== 200 ||
      !publicPages.has(url.pathname) ||
      !/^text\/html(?:\s*;|$)/i.test(contentType)
    ) {
      return upstream;
    }

    return new HTMLRewriter()
      .on('link[rel="canonical"]', {
        element(tag) {
          const href = tag.getAttribute('href');
          if (href === previewOrigin || href?.startsWith(`${previewOrigin}/`)) {
            tag.setAttribute('href', productionOrigin + href.slice(previewOrigin.length));
          }
        },
      })
      .on('meta[name="robots"]', {
        element(tag) {
          const content = tag.getAttribute('content')?.trim() ?? '';
          if (/^noindex\s*,\s*nofollow$/i.test(content)) {
            tag.remove();
          }
        },
      })
      .transform(upstream);
  },
};
