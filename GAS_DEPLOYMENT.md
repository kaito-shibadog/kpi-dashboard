# GAS デプロイ管理

## 現在のデプロイ状況

### ✅ 最新デプロイ

**GAS_URL**: 
```
https://script.google.com/macros/s/AKfycbw5Oqel437ID97nr5JZGVzE0dYvxpfDYqwVNkcmT7yKGyftqhL6bSxFEnFp-3RIDQ70mA/exec
```

**デプロイ日時**: 2026-08-06

**修正内容**:
- 業種別支援人数：`B4:M19` → `B4:N19` (10月データ追加)
- 業種別紹介人数：`B24:M39` → `B24:N39` (10月データ追加)

**検証済み**: ✅ JSON データで 12ヶ月分のデータが正しく取得されている

---

## バックアップコード

万が一 HTML の修正中にバグが発生した場合は、以下の手順で復帰できます：

1. `gas-backup/latest-gas-code.gs` をコピー
2. Google Apps Script エディタにペースト
3. 「新しいデプロイ」をクリック
4. 新しい URL を取得して `index.html` と `debug.html` に設定

---

## 履歴

| URL | 修正内容 | 状態 |
|-----|---------|------|
| AKfycbw5Oqel... | 業種別 N列 追加 | ✅ 現在使用中 |
| AKfycbxFui01... | 初期実装 | ⚠️ 10月データ欠落 |
