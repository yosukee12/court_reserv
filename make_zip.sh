ZIP_PATH="/Users/yosuke/develop/chatgpt/court_reserv.zip"

rm -f "$ZIP_PATH"

zip -r "$ZIP_PATH" . \
  -x "node_modules/" \
  -x "node_modules/*" \
  -x "*/node_modules/*" \
  -x "tmp_plan_*/*" \
  -x "*/tmp_plan_*/*" \
  -x "dist/" \
  -x "dist/*" \
  -x "*/dist/*" \
  -x ".terraform/*" \
  -x "*/.terraform/*" \
  -x ".git/*" \
  -x "*/.git/*" \
  -x ".DS_Store" \
  -x "*/.DS_Store" \
  -x "terraform.tfstate*" \
  -x "*/terraform.tfstate*" \
  -x ".env" \
  -x ".env.*" \
  -x "*/.env" \
  -x "*/.env.*" \
  -x "local.settings.json" \
  -x "*/local.settings.json" \
  -x "terraform.tfvars" \
  -x "*/terraform.tfvars" \
  -x "backend.tf" \
  -x "*/backend.tf" \
  -x "make_zip.sh"
