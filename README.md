# PowerShell創建捷徑

快速使用

```ps1
irm bit.ly/CreateShortcut|iex; CreateShortcut "README.md"
```

詳細說明

```ps1
# 載入函式庫
irm bit.ly/CreateShortcut|iex

# 創建捷徑到當前目錄
CreateShortcut "README.md"
# 創建捷徑到指定完整路徑 (-Path漏打副檔名會變成下一個的創建到指定目錄)
CreateShortcut "README.md" -Path "README.ink"
# 創建捷徑到指定目錄 (實際路徑為 .\DirName\README.md)
CreateShortcut "README.md" -Path "DirName"

# 捷徑的參數
CreateShortcut "README.md" -Arguments "Arguments"
# 捷徑的開始位置 (預設是取來源檔案的所在目錄)
CreateShortcut "README.md" -WorkingDirectory "C:\"
# 捷徑的描述
CreateShortcut "README.md" -Description "Test Description"
```
