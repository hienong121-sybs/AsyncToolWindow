# Hướng dẫn: Document & File APIs

**Tính năng:** Document & File APIs (Section 4)  
**Ngày:** 2026-03-19  
**Phiên bản:** patch `feat: document-file-apis`

---

## Mục tiêu

Cung cấp `DocumentService` – wrapper gọn, type-safe bao bọc DTE Document APIs – cho phép:
- Đọc thông tin file đang mở (Name, Path, Language, Saved, ReadOnly, Encoding)
- Liệt kê tất cả documents đang mở
- Thao tác tài liệu: Save, SaveAs, Close, Undo
- Mở file / điều hướng URL
- Đọc/ghi qua `TextDocument` + `EditPoint`
- Chạy built-in VS commands (`Edit.FormatDocument`, `File.SaveAll`, `Edit.GoToLine`, ...)

---

## File mới / thay đổi

| File | Loại thay đổi |
|------|---------------|
| `src/Services/DocumentService.cs` | **Mới** – service chính |
| `src/ToolWindows/SampleToolWindowState.cs` | Thêm property `Document` |
| `src/MyPackage.cs` | Khởi tạo và truyền `DocumentService` |
| `src/ToolWindows/SampleToolWindowControl.xaml` | Thêm 8 button mới |
| `src/ToolWindows/SampleToolWindowControl.xaml.cs` | Thêm 8 handler |
| `src/AsyncToolWindowSample.csproj` | Thêm compile entry `DocumentService.cs` |

---

## Cách sử dụng DocumentService

### 1. Khởi tạo (tự động qua MyPackage)

`DocumentService` được construct trong `MyPackage.InitializeAsync()` và truyền xuống `SampleToolWindowState.Document`. Không cần `InitializeAsync()` riêng vì service resolve DTE on-demand.

### 2. Đọc thông tin document

```csharp
ThreadHelper.ThrowIfNotOnUIThread();
var info = Document.GetActiveDocumentInfo();
// info.Name, .FullName, .Language, .Saved, .ReadOnly, .Encoding, .ProjectName
```

### 3. Liệt kê tất cả documents đang mở

```csharp
ThreadHelper.ThrowIfNotOnUIThread();
var docs = Document.GetAllOpenDocuments();
foreach (var d in docs)
    Console.WriteLine($"[{(d.Saved ? "✓" : "*")}] {d.Language} {d.Name}");
```

### 4. Lưu / đóng / undo

```csharp
Document.SaveActiveDocument();
Document.SaveActiveDocumentAs(@"C:\backup.cs");
Document.CloseActiveDocument(saveChanges: true);
Document.UndoActiveDocument();
```

### 5. Mở file / URL

```csharp
Document.OpenFile(@"C:\project\Program.cs");
Document.NavigateUrl("https://docs.microsoft.com");
```

### 6. TextDocument + EditPoint

```csharp
var tdInfo = Document.GetTextDocumentInfo();
// tdInfo.LastLine, .TotalChars, .Preview

string lines = Document.ReadLines(1, 6); // dòng 1–5

Document.InsertAtStart("// Auto-generated header\n");
```

### 7. ExecuteCommand

```csharp
Document.FormatDocument();              // Edit.FormatDocument
Document.SaveAll();                     // File.SaveAll
Document.GoToLine(42);                  // Edit.GoToLine 42
Document.ExecuteCommand("Build.BuildSolution");
```

---

## Buttons trong Tool Window

| Button | API gọi | Output |
|--------|---------|--------|
| Show Active Doc Info | `GetActiveDocumentInfo()` | Name/Path/Language/Saved/Encoding |
| List All Open Docs | `GetAllOpenDocuments()` | Danh sách [✓/*] Language Name |
| TextDoc Info + Preview | `GetTextDocumentInfo()` | Lines, Chars, 200-char preview |
| Read Lines 1–5 | `ReadLines(1, 6)` | Text của 5 dòng đầu |
| Save Active Document | `SaveActiveDocument()` | Lưu file hiện tại |
| Format Document | `FormatDocument()` | Chạy Edit.FormatDocument |
| Save All | `SaveAll()` | Chạy File.SaveAll |
| Go To Line 1 | `GoToLine(1)` | Nhảy về dòng 1 |

---

## Lưu ý thread safety

Tất cả public method của `DocumentService` đều yêu cầu UI thread. Luôn đảm bảo:

```csharp
// Trong button click handler (đã trên UI thread):
ThreadHelper.ThrowIfNotOnUIThread();

// Hoặc từ background:
await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
```

---

## DTOs

### `DocumentInfo`
| Property | Type | Mô tả |
|----------|------|-------|
| `Name` | string | Tên file (không path) |
| `FullName` | string | Đường dẫn đầy đủ |
| `Language` | string | "CSharp", "HTML", "Plain Text"... |
| `Kind` | string | GUID loại document |
| `Saved` | bool | false = có thay đổi chưa lưu |
| `ReadOnly` | bool | File read-only |
| `Encoding` | int | Windows CodePage (65001=UTF-8) |
| `ProjectName` | string? | Tên project chứa file |
| `ProjectFilePath` | string? | Đường dẫn file trong project |

### `TextDocumentInfo`
| Property | Type | Mô tả |
|----------|------|-------|
| `FirstLine` | int | Luôn = 1 |
| `LastLine` | int | Tổng số dòng |
| `TotalChars` | int | Tổng ký tự |
| `Preview` | string | 200 ký tự đầu tiên |
