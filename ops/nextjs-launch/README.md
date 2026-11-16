# Dental Circle Group launch handoff

This folder records publicly observable launch behavior and the remaining checks for the Next.js/Payload website. The Worker is the code supplied for deployment in the launch conversation.

## Verified on November 16, 2026

- The public HTTPS home page, Dental Implants page, and appointment form returned HTTP 200. Their canonical links pointed to the public site and the checked responses contained no noindex directive.
- HTTP on the bare domain and HTTPS on www redirected to the bare HTTPS site with status 301, preserving the tested path and query string.
- Public robots.txt points to sitemap.xml. The sitemap contains 23 public URLs.
- The appointment confirmation page returned HTTP 200 and retained noindex.
- Two legacy URLs were verified during launch. /about-us/ leads to /about and /dental-implants/ leads to /services/dental-implants.

## Remaining work in order

1. Submit one clearly marked test request with test data through https://dentalcirclegroup.com/appointment. Verify automatic navigation to /appointment/confirmation and verify that the same request appears once in Payload → Appointments. The staging submission reached Payload. The production submission remains unverified.
2. Compare the old URL inventory against the public site. Verify redirects for important legacy URLs and map missing destinations to equivalent pages where available.
3. Check that all sitemap URLs return 200 with matching production canonicals and no indexing blocks. Submit sitemap.xml to Google Search Console and inspect the homepage and priority services.
4. Locate the editable Next.js/Payload source. The launch conversation documents the temporary SEO fix. Once the source is recovered, implement and test canonical, robots and sitemap behavior in the app. Keep staging and the confirmation page noindex. Verify the replacement release before retiring the temporary Worker.

The production application source is currently unavailable.
