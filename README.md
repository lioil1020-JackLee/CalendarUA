# CalendarUA

CalendarUA 是一套以 PySide6 開發的 OPC UA 排程桌面工具，提供 Day / Week / Month 視圖、RRULE 週期規則、假日與例外處理，以及 SQLite 本地資料庫管理。

## 功能重點

- 行事曆視圖：`Day`、`Week`、`Month`。
- 同時段多事件：日/週格子可顯示並挑選多筆重疊事件進行編輯。
- 月視圖互動：雙擊事件 chip 進入編輯；雙擊空白日期新增。
- 假日與例外：可取消或覆寫單次 occurrence。
- OPC UA 寫值：支援 URL / NodeId / 值 / 資料型別配置。
- 顏色策略：目前事件統一紅底白字（已移除分類管理）。

## 專案結構

```text
CalendarUA/
├─ CalendarUA.py
├─ core/
│  ├─ lunar_calendar.py
│  ├─ opc_handler.py
│  ├─ rrule_parser.py
│  ├─ schedule_models.py
│  └─ schedule_resolver.py
├─ database/
│  └─ sqlite_manager.py
├─ ui/
│  ├─ database_settings_dialog.py
│  ├─ month_grid.py
│  ├─ recurrence_dialog.py
│  └─ schedule_canvas.py
├─ requirements.txt
├─ pyproject.toml
├─ CalendarUA-onedir.spec
└─ CalendarUA-onefile.spec
```

## 環境需求

- Python 3.9+
- Windows / Linux / macOS

## 開發安裝（建議使用 uv）

```bash
uv sync --dev
```

若你使用既有虛擬環境，也可用：

```bash
pip install -r requirements.txt
```

## 執行

```bash
uv run python CalendarUA.py
```

## 打包

### onedir

```bash
uv run pyinstaller --clean --noconfirm CalendarUA-onedir.spec
```

輸出：`dist/CalendarUA-onedir/`

### onefile

```bash
uv run pyinstaller --clean --noconfirm CalendarUA-onefile.spec
```

輸出：`dist/CalendarUA-onefile.exe`（Windows）

## 資料庫

- 預設使用 SQLite。
- 可在 `File -> Database Settings` 變更資料庫路徑、備份與還原。

## RRULE 說明

詳見 `rrule.md`。CalendarUA 使用 RFC 5545 RRULE，並支援自訂 `DURATION` 參數。

## 注意事項

- 若 OPC UA 節點需安全連線，請在排程設定中填入對應安全參數。
- 若打包後啟動失敗，請先以 `uv run python CalendarUA.py` 驗證開發環境可正常執行。
| node_id | TEXT | OPC UA 節點 ID | NOT NULL |
| target_value | TEXT | 目標數值 | NOT NULL |
| rrule_str | TEXT | RRULE 規則字串 | NOT NULL |
| opc_security_policy | TEXT | 安全策略 | DEFAULT 'None' |
| opc_security_mode | TEXT | 安全模式 | DEFAULT 'None' |
| opc_username | TEXT | 使用者名稱 |  |
| opc_password | TEXT | 密碼 |  |
| opc_timeout | INTEGER | 連線超時(秒) | DEFAULT 5 |
| opc_write_timeout | INTEGER | 寫值重試延遲(秒) | DEFAULT 3 |
| is_enabled | INTEGER | 啟用狀態 | DEFAULT 1 |
| created_at | TIMESTAMP | 建立時間 | DEFAULT CURRENT_TIMESTAMP |
| updated_at | TIMESTAMP | 更新時間 | DEFAULT CURRENT_TIMESTAMP |

## 🔍 故障排除

### 常見問題

**Q: 應用程式啟動失敗**
A: 確認 Python 版本 ≥ 3.9，並已安裝所有依賴套件

**Q: OPC UA 連線失敗**
A: 檢查伺服器位址、網路連線和安全設定。連線逾時預設為 5 秒，可在任務設定中調整

**Q: 排程未執行**
A: 確認任務已啟用，且 RRULE 規則設定正確

**Q: 寫值失敗**
A: 檢查 Node ID 是否正確，寫值重試延遲預設為 3 秒，可在任務設定中調整

**Q: 資料庫錯誤**
A: 檢查資料庫檔案權限，或嘗試重新建立資料庫

### 記錄與除錯

系統預設記錄等級為 WARNING。如需詳細記錄，請修改 `CalendarUA.py` 中的記錄設定：

```python
logging.basicConfig(level=logging.DEBUG)
```

## 🤝 貢獻指南

歡迎參與專案開發！請遵循以下步驟：

1. Fork 此專案
2. 建立功能分支 (`git checkout -b feature/AmazingFeature`)
3. 提交變更 (`git commit -m 'Add some AmazingFeature'`)
4. 推送至分支 (`git push origin feature/AmazingFeature`)
5. 開啟 Pull Request

### 開發環境設定

```bash
# 安裝開發依賴
pip install -e ".[dev]"

# 執行測試
python -m pytest

# 建置可執行檔案
pyinstaller CalendarUA.spec
```

## 📄 授權條款

本專案採用 MIT 授權條款 - 詳見 [LICENSE](LICENSE) 檔案

## 📞 聯絡資訊

- **專案維護者**: [lioil1020-JackLee](https://github.com/lioil1020-JackLee)
- **問題回報**: [GitHub Issues](https://github.com/lioil1020-JackLee/CalendarUA/issues)
- **專案首頁**: [GitHub Repository](https://github.com/lioil1020-JackLee/CalendarUA)

---

**注意**: 本系統適用於工業自動化環境，請在專業人員指導下使用。如有特殊需求或客製化需求，請聯絡開發團隊。