# SharePoint Cleanup API 利用手順

## 1. 依存関係インストール

```powershell
python -m pip install -r requirements.txt
```

## 2. 環境変数設定

`env.example` を参考に `.env` 相当を設定してください。

必須:

- `SP_TENANT_ID`
- `SP_CLIENT_ID`
- `SP_CERT_THUMBPRINT`
- `SP_PRIVATE_KEY_PATH`
- `SP_SITE_URL`
- `SP_ROOT_FOLDER_SERVER_RELATIVE_URL`

推奨:

- `SP_API_KEY`

## 3. APIサーバ起動

```powershell
python sharepoint_version_cleanup.py --serve --host 0.0.0.0 --port 8000
```

## 4. エンドポイント

- `GET /health`
- `POST /cleanup`（同期）
- `POST /cleanup/background`（非同期）
- `GET /cleanup/background/{job_id}`（非同期ジョブ状態確認）

## 5. リクエスト例

```json
{
  "root_folder_server_relative_url": "/sites/example/Shared Documents",
  "days_to_keep": 30,
  "dry_run": true
}
```

HTTPヘッダー:

- `X-API-Key: <SP_API_KEY>`
