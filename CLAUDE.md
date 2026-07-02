# CLAUDE.md — Susi 媽媽血壓記錄 LINE Bot（Google Apps Script）

本 repo 是「AI Care Ops Assistant」產品的一個模組（生命徵象核心循環的 LINE 輸入端）。
完整產品脈絡——媽媽的臨床限制、看護 persona、15 秒操作鐵律、警報降噪原則、
投資人就緒優先順序——定義在上層目錄 `../CLAUDE.md`，動工前先讀。

## 本 repo 的特殊地位

[`blood_pressure_bot_docs.md`](blood_pressure_bot_docs.md) 裡的血壓分級判定表
（客製化給 82 歲、1.4 cm 未破裂腦動脈瘤、骨鬆、曾低血壓的長輩）是整個產品
**血壓警示邏輯的 single source of truth**。其他 repo（如 `../blood-pressure-tracker`）
的警示門檻必須與它一致。改門檻前必須明確告知使用者。

## 已知的坑

- GAS 時區 Bug：歷史紀錄與重複比對一律用 `getDisplayValues()`，
  不要用 `getValues()`（詳見 docs 第 2 節）。
- 時段判定直接解析台北時間字串，不依賴伺服器時區。
