import fs from "node:fs";

for (const key of ["MS_TENANT_ID", "MS_CLIENT_ID"]) {
  if (!process.env[key]) throw new Error(`Missing required environment variable: ${key}`);
}

const scope = "offline_access Files.ReadWrite User.Read Mail.Send";
const response = await fetch(`https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/devicecode`, {
  method: "POST",
  headers: { "Content-Type": "application/x-www-form-urlencoded" },
  body: new URLSearchParams({ client_id: process.env.MS_CLIENT_ID, scope })
});
if (!response.ok) throw new Error(`Could not start Microsoft authorization (${response.status}): ${await response.text()}`);

const device = await response.json();
fs.writeFileSync("/tmp/email-device-code.json", JSON.stringify(device));
console.log("\nMICROSOFT AUTHORIZATION REQUIRED");
console.log(device.message);
console.log("Keep this workflow running while authorization is completed.\n");

if (process.env.GITHUB_STEP_SUMMARY) {
  fs.appendFileSync(
    process.env.GITHUB_STEP_SUMMARY,
    `## Microsoft authorization required\n\nOpen [microsoft.com/devicelogin](${device.verification_uri}) and enter code **${device.user_code}**.\n\nKeep this workflow running until authorization completes.\n`
  );
}
