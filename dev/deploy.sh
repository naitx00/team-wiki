#!/usr/bin/env bash
# Deploys the wiki to its live link: push, then a new version of the existing
# deployment (the /exec URL stays the same).
#
#   bash dev/deploy.sh "what changed"
#
# Every deployment must be "Who has access: Anyone within your domain" and
# "Execute as: Me". Those come from appsscript.json, which every push uploads,
# so this refuses to deploy if the file says anything else.
set -euo pipefail
cd "$(dirname "$0")/.."

DEPLOYMENT_ID="AKfycbwkv6Sm6kVxcyr7A8Eor_PFn6br33oJFOatj3gXClycGc4wJgthRzOCXBttXtUwL90f"
DESCRIPTION="${1:-Update}"

node -e '
  const m = JSON.parse(require("fs").readFileSync("appsscript.json", "utf8"));
  const w = m.webapp || {};
  const problems = [];
  if (w.access !== "DOMAIN") problems.push("access is " + JSON.stringify(w.access) + ", must be \"DOMAIN\" (Anyone within your domain)");
  if (w.executeAs !== "USER_DEPLOYING") problems.push("executeAs is " + JSON.stringify(w.executeAs) + ", must be \"USER_DEPLOYING\" (Execute as: Me)");
  if (problems.length) { console.error("Not deploying — appsscript.json: " + problems.join("; ")); process.exit(1); }
  console.log("appsscript.json OK: Anyone within your domain, execute as me");
'

node dev/test-server.js > /dev/null && echo "Server tests passed"

# --force: Git Bash can't show clasp's "overwrite the manifest?" prompt.
clasp push --force
clasp create-deployment --deploymentId "$DEPLOYMENT_ID" --description "$DESCRIPTION"
