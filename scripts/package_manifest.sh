#!/usr/bin/env bash
# teams-package/ の manifest.json + アイコン2点を zip 化してサイドロード用パッケージを作る
set -euo pipefail

cd "$(dirname "$0")/../teams-package"
rm -f teams-app-package.zip
zip -r teams-app-package.zip manifest.json color.png outline.png
echo "作成しました: teams-package/teams-app-package.zip"
