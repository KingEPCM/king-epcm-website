# King EPCM — SEO & GEO Review

_Reviewed July 1, 2026 · site: `king-epcm-site/` → kingepcm.com_

## ✅ Completed in this session (on-site edits)

- **Legacy Wix 301 redirects** added to `staticwebapp.config.json` (e.g. `/geotechnical` → `/geotechnical-hydrogeology`, plus 15 common old slugs).
- **Homepage business schema enriched**: geo coordinates (approx — verify), business hours, `priceRange`, payment methods, and explicit city/region `areaServed`.
- **Team page**: added `Person`/`employee` schema for all 12 team members (E-E-A-T).
- **Careers page**: added `WebPage` + `BreadcrumbList` schema (JobPosting deliberately not faked — add per real opening).
- **Services hub**: added full `OfferCatalog` of the 7 services.
- **Service page titles**: added Markham/GTA/Ontario/Canada geo keywords to all 7.
- **Sitemap**: added `<lastmod>` dates to every URL.
- **404 page**: created a branded `404.html` and pointed the config's 404 handler to it (was showing the homepage).
- **Meta descriptions**: trimmed the 5 that exceeded 160 chars (contact, mining-aggregate, project-onboarding, services, team).
- `llms.txt` already existed and is good — left as-is.

**Still off-site / needs you:** legacy URL list from Search Console, NAP cleanup in directories, Google Business Profile, collecting reviews (+ Review schema), and a blog/Insights section. Details below.

---


## Bottom line

The site is already in strong technical shape. Every page has a unique title, meta description, single H1, canonical tag, Open Graph + Twitter cards, and clean URLs with 301 redirects. Structured data is rich — all seven service pages carry `Service` + `ProfessionalService` + `BreadcrumbList` + `FAQPage` schema, the homepage has an `OfferCatalog` + `FAQPage`, and every image has alt text. This is better than most firms in this sector.

The gaps are mostly **off-page (local citations, Google Business Profile), a few schema additions, and a content-marketing hole** that limits both classic SEO and AI-answer visibility (GEO).

---

## High priority

**1. Legacy Wix URL redirects (fixes duplicate content + 404s).**
Google still indexes old Wix paths — e.g. `kingepcm.com/geotechnical` and `kingepcm.com/contact` — which don't exist on the new site (the new page is `/geotechnical-hydrogeology`). Add 301 redirects in `staticwebapp.config.json` from every old Wix path to its new equivalent so ranking signals transfer and visitors don't hit dead ends.

**2. NAP (Name/Address/Phone) consistency across the web.**
External directories are inconsistent — a public geotechnical report and some listings show `647-459-5647`, while the site uses `416-342-3001`. Inconsistent NAP directly hurts local ranking. Audit and correct Google Business Profile, Yellowpages, RocketReach, ZoomInfo, LinkedIn so all show the same phone and address. On-site, standardize `3780 14th Ave` vs `3780 14th Avenue` (33 pages use "Avenue", 15 use "Ave") — pick one, ideally the full "3780 14th Avenue, Unit 211".

**3. Google Business Profile.**
The single biggest lever for a local engineering firm. Ensure it's claimed, category = "Engineer / Civil engineer / Environmental consultant", office photos added, and actively collecting reviews. (There's already a `GOOGLE_BUSINESS_PROFILE_CHECKLIST.md` in the parent folder — work through it.)

**4. Add client reviews + testimonials with `Review`/`aggregateRating` schema.**
No review markup exists anywhere. Real testimonials with schema are a strong trust signal for Google and a heavily-weighted citation signal for AI engines.

---

## Medium priority

**5. Add an Insights / Resources section (blog).** _Biggest content gap._
There is no blog or article section. This limits long-tail keyword capture, content freshness, and — critically for GEO — gives AI engines nothing original to cite. Publish practical, question-shaped articles that match how clients actually search and how AI fans out queries, e.g.:
- "Do I need a geotechnical report for my Ontario building permit?"
- "Phase I vs Phase II ESA — what's the difference and when do you need each?"
- "What is a functional servicing report and who requires it?"
Answer the question directly in the first two sentences, then expand. This format wins both featured snippets and AI citations.

**6. Add `Person` schema to the Team page.**
`team.html` has zero `Person` markup. Adding it (name, jobTitle, credentials) builds the expertise/authority signal (E-E-A-T) that both Google and AI models use to judge trustworthiness.

**7. Add `JobPosting` schema to Careers.**
`careers.html` has no structured data. `JobPosting` schema makes openings eligible for the Google Jobs rich result — free, high-intent traffic.

**8. Enrich the Services hub page.**
`services.html` only has `BreadcrumbList`. Add an `OfferCatalog`/`ItemList` of the seven services (like the homepage) so the overview page is machine-readable too.

**9. Add location keywords to service-page titles.**
Detail-page titles are generic — e.g. `Chemical Environmental | King EPCM`. Add the geo modifier: `Chemical Environmental (Phase I & II ESA) | Markham & GTA | King EPCM`. The meta descriptions already mention Markham/GTA; the titles should too, since titles carry more ranking weight.

---

## Local SEO (geo) schema additions

**10. Enrich the `ProfessionalService` schema.** It currently has address, phone, email, logo, `sameAs` (LinkedIn + Facebook), and `foundingDate`, but is missing:
- `geo` with `latitude`/`longitude` (helps map/local pack placement)
- `openingHours` (business hours)
- `areaServed` as an explicit list of cities/regions (Markham, Toronto, GTA, Golden Horseshoe) rather than prose
- `priceRange` (even "$$" satisfies Google's expected field)

---

## GEO (AI answer engines) — specific notes

The FAQ schema across the site is exactly what GEO rewards, so the foundation is good. To go further:

**11. Lead every FAQ answer with the direct answer** in the first sentence (AI extracts the opening, not the build-up). Most are close; review for any that warm up before answering.

**12. Add `<lastmod>` dates to `sitemap.xml`.** It currently has none. AI engines weigh recency; dated content signals freshness.

**13. Consider an `llms.txt` file** at the site root — an emerging GEO practice: a short plain-text summary of the firm, services, and service area that AI crawlers read to understand the business quickly.

**14. Confirm AI crawlers aren't blocked.** `robots.txt` currently allows all bots (good). If the domain ever moves behind Cloudflare, re-check — Cloudflare now blocks AI bots by default, which would silently remove the site from AI answers.

---

## Low priority / polish

**15. Dedicated 404 page.** `staticwebapp.config.json` rewrites unknown URLs to the homepage with a 404 status. A purpose-built 404 with helpful links is better UX (and cleaner for crawlers).

**16. Trim a few long meta descriptions** so they don't truncate in results: team (190 chars), services (184), homepage (174), and several service pages (167–171). Target ~150–155.

**17. Standardize email display casing** (`info@KingEPCM.com` vs `info@kingepcm.com`) — cosmetic only (email is case-insensitive), but consistency looks sharper.

---

## Suggested order of attack

1. Legacy URL 301s + NAP cleanup + Google Business Profile (fastest ranking wins)
2. Reviews/testimonials + `Person`/`JobPosting`/local-schema additions
3. Launch an Insights section and publish 3–5 question-based articles
4. GEO polish: sitemap lastmod, llms.txt, FAQ answer-first pass

Items 1, 8, 9, 10, 12, 13, 15, 16 are all edits I can make directly in these files. Items 2, 3, 4 (directories, GBP, collecting reviews) need to happen off-site.
