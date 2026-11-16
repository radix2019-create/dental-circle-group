# Dental Circle Group launch handoff

This folder records publicly observable launch behavior and the remaining checks for the Next.js/Payload website. The Worker is the code supplied for deployment in the launch conversation.

## Verified on November 16, 2026

- The public HTTPS home page, Dental Implants page, and appointment form returned HTTP 200. Their canonical links pointed to the public site and the checked responses contained no noindex directive.
- HTTP on the bare domain and HTTPS on www redirected to the bare HTTPS site with status 301, preserving the tested path and query string.
- Public robots.txt points to sitemap.xml. All 23 sitemap URLs returned HTTP 200 with matching production canonicals and no noindex in the checked responses.
- The appointment confirmation page displayed the success message on the public domain and retained noindex. The launch test request appeared in the appointment dashboard.
- Verified legacy redirects for /about-us/, /dental-implants/, /root-canal-therapy/, and /preventative-dentistry/.
- Verified 12 new legacy redirects for orthodontics, periodontics, cosmetic dentistry, oral surgery, Invisalign, and veneers, with and without a trailing slash. Each returned one 301 to the matching /services/ page, preserved the test query string, and ended with HTTP 200.

- Google Search Console read the production sitemap successfully on November 16, 2026, and reported 23 discovered pages.

## Remaining work in order

1. Inspect the live homepage and priority services in Google Search Console.
2. Continue auditing legacy URLs. Identify equivalent destinations for /new-patients/ and /pediatric-dentistry/ before adding redirects.
3. Locate the editable Next.js/Payload source. The launch conversation documents the temporary SEO fix. Once the source is recovered, implement and test canonical, robots and sitemap behavior in the app. Keep staging and the confirmation page noindex. Verify the replacement release before retiring the temporary Worker.

The production application source is currently unavailable.
