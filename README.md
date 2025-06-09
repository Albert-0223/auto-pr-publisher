# 🔧 GitHub Auto PR Creator

A Python automation tool for creating GitHub Pull Requests in bulk using compare links provided in Excel files. Ideal for internal publishing workflows.

---

## 🚀 Features

- ✅ **Excel-driven input**: Read GitHub compare links from `.xlsx` files
- ✅ **Edge integration**: Open pre-check and PR links in Microsoft Edge with grouped tabs
- ✅ **Timezone-aware PR titles**: PR title auto-generates based on current PST/PDT time with AM/PM
- ✅ **Smart filtering**:
  - Skip empty compare links (no commits)
  - Handle duplicate PRs by checking if already exists
  - Catch and classify errors (invalid repo, token issues, etc.)
- ✅ **Detailed Excel output**: Summarizes results with auto-formatting
- ✅ **Robust logging**: Each run saves logs with timestamped filenames
- ✅ **Parallel processing**: Handles links concurrently with threading

---

## 📁 Input Files

- `INPUT_PATH`: Excel file with **compare links** in the first column (from row 2)
- `PRE_LINKS_FILE`: CSV/XLSX with optional links to open in browser before PR creation

### 🔗 Compare Link Format
Each compare link should look like:
```
https://github.com/{org}/{repo}/compare/{base}...{head}
```
Example:
```
https://github.com/my-org/my-repo/compare/dev...main
```

---

## 📦 Installation & Dependencies

### 🐍 Python Version
Python 3.10 or later is required.

### 📦 Install Required Packages
```bash
pip install openpyxl requests
```

---

## 🛠 Configuration

At the top of the script, configure:

```python
GITHUB_TOKEN = "ghp_xxx..."  # GitHub personal access token
INPUT_PATH = r"path/to/compare_links.xlsx"
PRE_LINKS_FILE = r"path/to/pre_check_links.csv"
BASE_OUTPUT_DIR = os.path.join(os.path.expanduser("~"), "Desktop", "PR_created_result")
```

> You can switch among multiple pre-defined files by uncommenting the desired path.

---

## 📋 Excel Output Preview

The script generates an Excel report after each run:

| Compare Link | Result   | PR Link                                     | Commits | Files Changed | Reason                      |
|--------------|----------|----------------------------------------------|---------|----------------|-----------------------------|
| https://github.com/org/repo/compare/dev...main | Created  | https://github.com/org/repo/pull/123 | 5       | 12             |                             |
| https://github.com/org/repo/compare/feat...main | Skipped  | -                                            | 0       | 8              | No new commits to publish. |
| https://github.com/org/repo/compare/bug...main | Duplicate| https://github.com/org/repo/pull/119 | 3       | 300+           | Pull request already exists. |
| https://github.com/org/repo/compare/exp...main | Error    | -                                            | -       | -              | Repository or branch not found. |

- `300+` indicates GitHub API truncated file list at 300
- `-` means no PR was created
- `Reason` explains the result

---

## 🖥 How to Use

### ✅ Step-by-step
1. Ensure you’ve installed dependencies
2. Update `INPUT_PATH`, `PRE_LINKS_FILE`, and `GITHUB_TOKEN`
3. Run:
```bash
python main.py
```
4. Confirm prompts as needed
5. Output files and logs will appear in your desktop directory

### 🧪 Example Prompt
```bash
📄 File to process: C:\...\OPS-Publish-10_00.xlsx
⚠️ Confirm to start PR creation for this file? (y/n): y
```

---

## 🌐 Browser Integration

- Pre-check links and created PRs are opened using Microsoft Edge
- Links are grouped by tab windows to prevent overload

---

## 🧠 Notes

- Ensure your GitHub token has **repo** and **pull_request** scopes
- If Excel output fails (due to open file), script saves a fallback `*_retry.xlsx`
- Edge browser path is hardcoded as:
```python
C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe
```

Update if needed.

---

## 🧰 Advanced Customization

- You can adjust threading level by modifying:
```python
ThreadPoolExecutor(max_workers=5)
```
- Default Excel column formatting can be tweaked in `save_results_to_excel()`
- Grouping size for Edge tab batches can be tuned in:
```python
def open_links_in_edge_window_grouped(links, group_size=15)
```

---

## 📬 Feedback

Created by internal tooling team. Contact Bowen for improvements or token config support.

