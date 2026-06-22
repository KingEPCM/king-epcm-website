/* Diagnostic endpoint: GET /api/ping
 * Confirms the API is deployed and reports the runtime + whether the libraries
 * the onboarding function needs are actually available. Safe to leave or delete. */
module.exports = async function (context, req) {
  let pdfkit = "not loaded";
  try { require("pdfkit"); pdfkit = "loaded OK"; }
  catch (e) { pdfkit = "FAILED: " + ((e && e.message) || e); }

  context.res = {
    status: 200,
    headers: { "Content-Type": "application/json" },
    body: {
      ok: true,
      marker: "ping-v1",
      node: process.version,
      hasFetch: typeof fetch !== "undefined",
      hasFormData: typeof FormData !== "undefined",
      hasBlob: typeof Blob !== "undefined",
      pdfkit: pdfkit,
      env: {
        AAD_CLIENT_ID: !!process.env.AAD_CLIENT_ID,
        AAD_CLIENT_SECRET: !!process.env.AAD_CLIENT_SECRET,
        ONBOARD_MAIL_FROM: process.env.ONBOARD_MAIL_FROM || null,
        ONBOARD_MAIL_TO: process.env.ONBOARD_MAIL_TO || null,
        TEAMWORK_DOMAIN: process.env.TEAMWORK_DOMAIN || null,
        TEAMWORK_API_KEY: !!process.env.TEAMWORK_API_KEY,
        TEAMWORK_INTAKE_PROJECT_ID: process.env.TEAMWORK_INTAKE_PROJECT_ID || null
      }
    }
  };
};
