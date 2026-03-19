# Changelog

All notable changes to **AsyncToolWindowSample** are documented here.
Format follows [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

---

## [Unreleased] – 2026-03-19 (feat: document-file-apis)

### Added

#### `src/Services/DocumentService.cs` *(new file)*
- **Mục đích:** Tách biệt logic truy cập Document & File APIs thành service riêng.
- **Public API:**
  - `GetActiveDocumentInfo()` – snapshot properties của active document (Name, FullName, Language, Kind, Saved, ReadOnly, Encoding, ProjectName, ProjectFilePath).
  - `GetAllOpenDocuments()` – danh sách tất cả documents đang mở.
  - `SaveActiveDocument()` – lưu document hiện tại.
  - `SaveActiveDocumentAs(path)` – save-as.
  - `CloseActiveDocument(saveChanges)` – đóng với tùy chọn lưu/bỏ.
  - `UndoActiveDocument()` – undo thay đổi cuối.
  - `OpenFile(path)` – mở file qua `ItemOperations.OpenFile`.
  - `NavigateUrl(url)` – điều hướng URL qua `ItemOperations.Navigate`.
  - `GetTextDocumentInfo()` – đọc FirstLine, LastLine, TotalChars, Preview 200 chars via `TextDocument` + `EditPoint`.
  - `ReadLines(startLine, endLine)` – đọc range dòng qua `EditPoint.GetLines`.
  - `InsertAtStart(text)` – chèn text ở đầu file qua `EditPoint.Insert`.
  - `ExecuteCommand(name, args)` – chạy built-in VS command.
  - `FormatDocument()` – `Edit.FormatDocument`.
  - `SaveAll()` – `File.SaveAll`.
  - `GoToLine(line)` – `Edit.GoToLine`.
- **DTOs:** `DocumentInfo`, `TextDocumentInfo`.
- **Lý do:** Demo Section 4 "Document & File APIs" theo tài liệu VS2017 Extension API Reference.

#### `docs/tutorials/document-file-apis_2026-03-19.md` *(new file)*
- Hướng dẫn sử dụng tính năng Document & File APIs: API reference, thread safety, DTOs, buttons.

### Changed

#### `src/ToolWindows/SampleToolWindowState.cs`
- **Thêm** property `DocumentService Document`.

#### `src/MyPackage.cs`
- **Thêm** property `DocumentService Document` (singleton).
- **`InitializeAsync`:** construct `DocumentService` (không cần `InitializeAsync` riêng).
- **`InitializeToolWindowAsync`:** populate `SampleToolWindowState.Document`.

#### `src/ToolWindows/SampleToolWindowControl.xaml`
- **Thêm** section "── Document & File APIs ──" với 8 button mới:
  - "Show Active Doc Info" – hiện Name/Path/Language/Saved/Encoding/Project.
  - "List All Open Docs" – liệt kê [✓/*] Language Name cho mọi document đang mở.
  - "TextDoc Info + Preview" – hiện Lines, TotalChars, 200-char preview qua TextDocument.
  - "Read Lines 1–5" – đọc và log 5 dòng đầu qua EditPoint.
  - "Save Active Document" – lưu file hiện tại.
  - "Format Document" – chạy Edit.FormatDocument.
  - "Save All" – chạy File.SaveAll.
  - "Go To Line 1" – nhảy về dòng 1 qua Edit.GoToLine.
- **`d:DesignHeight`** tăng từ 720 → 1080.

#### `src/ToolWindows/SampleToolWindowControl.xaml.cs`
- **Thêm** property `Document => _state.Document`.
- **Thêm** 8 handler tương ứng 8 button mới.

#### `src/AsyncToolWindowSample.csproj`
- **Thêm** `<Compile Include="Services\DocumentService.cs" />`.

#### `README.md`
- Cập nhật features table thêm Section 4 Document & File APIs.
- Cập nhật source map.

---

## [Unreleased] – 2026-03-19 (patch: CS0122-getservice)

### Fixed

#### `src/Services/SelectionService.cs`
- **CS0122 (×2) – `AsyncPackage.GetService(Type)` inaccessible:**
  `AsyncPackage.GetService()` là `protected internal` — không thể gọi từ class bên ngoài.
  **Fix:** Lưu `AsyncPackage` vào `IServiceProvider _serviceProvider` (interface public mà
  `AsyncPackage` implement). Thay tất cả `_package.GetService(...)` bằng
  `_serviceProvider.GetService(...)`.

---

## [Unreleased] – 2026-03-19 (patch: compiler-errors)

### Fixed

#### `src/Services/StatusBarService.cs`
- **CS0165 – Use of unassigned local variable 'frozen':** Fix init `frozen = 0`.
- **CS1503 (×2) – ref ulong → ref uint:** `IVsStatusbar.Progress` nhận `ref uint`.

#### `src/ToolWindows/SampleToolWindowControl.xaml.cs`
- Đổi `ulong cookie = 0` → `uint cookie = 0`.

---

## [Unreleased] – 2026-03-19 (feat: output-window-status-bar)

### Added

- `src/Services/OutputWindowService.cs` *(new)*
- `src/Services/StatusBarService.cs` *(new)*

### Changed

- `MyPackage.cs`, `SampleToolWindowState.cs`, `SampleToolWindowControl.xaml/.cs` – wire-up.

---

## [1.1] – baseline

- Async Tool Window cơ bản với button "Show VS Location".
- AsyncPackage load trên background thread.
- Command `ShowToolWindow` trong menu *View > Other Windows*.
