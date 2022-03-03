VBA > テーブル > 列
# get
## インデックス
1～  
```vba
Dim col As ListColumn
set col = シート.ListObjects(1).ListColumns(i)
```
## [列名](列名.md)
```vba
Dim col As ListColumn
set col = シート.ListObjects("列名").ListColumns(i)
```

## 全体
```vba
Dim col As ListColumns
set cols = シート.ListObjects(1).ListColumns
```
