#!/bin/bash
set -e
cd "$(dirname "$0")"

# Step 1: Restore Vite source template (index.html bị ghi đè sau mỗi deploy)
cat > index.html << 'VITE_TEMPLATE'
<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <link rel="icon" type="image/svg+xml" href="/vite.svg" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>Tra cứu tồn kho</title>
  </head>
  <body>
    <div id="root"></div>
    <script type="module" src="/src/main.jsx"></script>
  </body>
</html>
VITE_TEMPLATE

# Step 2: Build từ JSX source
node ./node_modules/vite/bin/vite.js build

# Step 3: Deploy - copy built files ra root (không dùng dist/)
rm -rf assets
cp -r dist/assets .
cp dist/index.html index.html

echo "✓ Build done → index.html + assets/ ready for GitHub Pages"
