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
# 創建捷徑到指定完整路徑(副檔名為非lnk時會自動補上lnk)
CreateShortcut "README.md" -Path "README.ink"
# 創建捷徑到當前目錄(捷徑名字自動取原檔名)
CreateShortcut "README.md" -Path "DirName"

# 捷徑的參數
CreateShortcut "README.md" -Arguments "Arguments"
# 捷徑的開始位置
CreateShortcut "README.md" -WorkingDirectory "C:\"
# 捷徑的描述
CreateShortcut "README.md" -Description "Test Description"
```
