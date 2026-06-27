/*
 * POST /api/drive-time  — driving time from the King EPCM office to a site address,
 * via Azure Maps (Search/Geocoding + Route Directions). Used by the Pre-Visit Checklist
 * to auto-fill the "Drive time (each way)" field.
 *
 * Request:  { "address": "1017 Victoria St N, Kitchener, ON" }
 * Response: { ok:true, minutes, text:"1 h 12 min", distanceKm:118, destination:"<matched address>" }
 *           { ok:false, error:"<message>" }
 *
 * Config (Static Web App application settings):
 *   AZURE_MAPS_KEY   (required)  — Azure Maps account subscription key
 *   OFFICE_ADDRESS   (optional)  — origin address; defaults to the Markham office
 */
const MAPS_KEY = process.env.AZURE_MAPS_KEY;
const OFFICE_ADDRESS = process.env.OFFICE_ADDRESS || "3780 14th Avenue, Markham, ON L3R 9Y5, Canada";

let _officeCoords = null;   // cached across warm invocations

module.exports = async function (context, req) {
  try {
    if (!MAPS_KEY) { return reply(context, { ok: false, error: "Drive-time isn't configured yet (no Azure Maps key set)." }); }
    const address = ((req.body && req.body.address) || "").trim();
    if (!address) { return reply(context, { ok: false, error: "Enter the site location first." }); }

    // 1) origin (office) — geocode once, then cache
    if (!_officeCoords) {
      _officeCoords = await geocode(OFFICE_ADDRESS);
      if (!_officeCoords) { return reply(context, { ok: false, error: "Could not locate the office address." }); }
    }
    // 2) destination (site)
    const dest = await geocode(address);
    if (!dest) { return reply(context, { ok: false, error: "Couldn't find that site address — check the spelling or add the city/province." }); }

    // 3) route office -> site
    const route = await driveRoute(_officeCoords, dest);
    if (!route) { return reply(context, { ok: false, error: "Couldn't calculate a driving route to that address." }); }

    return reply(context, {
      ok: true,
      minutes: Math.round(route.seconds / 60),
      text: fmtDur(route.seconds),
      distanceKm: Math.round(route.meters / 1000),
      destination: dest.label || ""
    });
  } catch (e) {
    context.log.error(e);
    return reply(context, { ok: false, error: "Drive-time service error. Please enter it manually." });
  }
};

/* ---------- Azure Maps ---------- */
async function geocode(query) {
  const url = "https://atlas.microsoft.com/search/address/json?api-version=1.0&limit=1&countrySet=CA"
    + "&subscription-key=" + encodeURIComponent(MAPS_KEY)
    + "&query=" + encodeURIComponent(query);
  const r = await fetch(url);
  if (!r.ok) return null;
  const j = await r.json();
  const hit = j && j.results && j.results[0];
  if (!hit || !hit.position) return null;
  return { lat: hit.position.lat, lon: hit.position.lon, label: (hit.address && hit.address.freeformAddress) || "" };
}

async function driveRoute(o, d) {
  const url = "https://atlas.microsoft.com/route/directions/json?api-version=1.0&travelMode=car&routeType=fastest"
    + "&subscription-key=" + encodeURIComponent(MAPS_KEY)
    + "&query=" + o.lat + "," + o.lon + ":" + d.lat + "," + d.lon;
  const r = await fetch(url);
  if (!r.ok) return null;
  const j = await r.json();
  const s = j && j.routes && j.routes[0] && j.routes[0].summary;
  if (!s) return null;
  return { seconds: s.travelTimeInSeconds, meters: s.lengthInMeters };
}

/* ---------- helpers ---------- */
function fmtDur(sec) {
  const m = Math.round(sec / 60);
  if (m < 60) return m + " min";
  const h = Math.floor(m / 60), mm = m % 60;
  return h + " h" + (mm ? " " + mm + " min" : "");
}
function reply(context, obj) {
  context.res = { status: 200, headers: { "Content-Type": "application/json" }, body: obj };
}
