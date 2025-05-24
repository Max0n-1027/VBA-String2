# String2 - VBA 文字列操作クラス

`String2` クラスは、VBA（Visual Basic for Applications）でJavaScriptのStringオブジェクトに類似した高機能な文字列操作を提供するユーティリティクラスです。

## 特徴

- JavaScriptライクなメソッド名と動作
- 正規表現の利用が可能（`VBScript.RegExp` を使用）
- 文字列の検索、分割、置換、トリミングなど、多数のユーティリティ関数を提供
- クラスのインスタンスを使ってメソッドチェーン可能

---

## クラスの使用方法

### インスタンスの作成

```vb
Dim str As String2
Set str = New String2
str.Value = "Hello, World"
````

または

```vb
Dim str As New String2 : str = "Hello, World"
```

---

## プロパティ

| 名前       | 型      | 説明                 |
| -------- | ------ | ------------------ |
| `Value`  | String | 現在の文字列値            |
| `length` | Long   | 文字列の長さ（バイトではなく文字数） |

---

## 主なメソッド一覧

### 基本操作

| メソッド                           | 戻り値型     | 説明           |
| ------------------------------ | -------- | ------------ |
| `NewInstance([Val As String])` | String2  | 新しいインスタンスを生成 |
| `ToString()`                   | String   | 文字列を取得       |
| `ToCharArray()`                | String() | 1文字ごとの配列を取得  |
| `ToLowerCase()`                | String2  | 小文字化         |
| `ToUpperCase()`                | String2  | 大文字化         |

---

### 文字列検索・抽出

| メソッド                                           | 戻り値型     | 説明                      |
| ---------------------------------------------- | -------- | ----------------------- |
| `CharAt(position As Long)`                     | String2  | 指定位置の1文字を取得             |
| `CharCodeAt(index As Long)`                    | Long     | 指定位置のUnicodeコードを取得      |
| `Includes(searchString As String, [start])`    | Boolean  | 部分文字列が含まれるか             |
| `IndexOf(searchString As String, [start])`     | Long     | 最初に出現する位置（なければ -1）      |
| `LastIndexOf(searchString As String, [start])` | Long     | 最後に出現する位置（なければ -1）      |
| `Search(pattern As String)`                    | Long     | 正規表現による検索。最初の一致位置または -1 |
| `Match(pattern As String)`                     | String() | 正規表現で一致した文字列の配列         |

---

### 文字列の変更・整形

| メソッド                                                                       | 戻り値型    | 説明                |
| -------------------------------------------------------------------------- | ------- | ----------------- |
| `Concat(Val As String)`                                                    | String2 | 文字列を結合            |
| `PadStart(targetLength As Long, [padString As String])`                    | String2 | 指定長さになるまで前方にパディング |
| `PadEnd(targetLength As Long, [padString As String])`                      | String2 | 指定長さになるまで後方にパディング |
| `Repeat(Count As Long)`                                                    | String2 | 指定回数繰り返し          |
| `Replace(pattern As String, replacement As String, [useRegex As Bool])`    | String2 | 最初の一致部分を置換（正規表現可） |
| `ReplaceAll(pattern As String, replacement As String, [useRegex As Bool])` | String2 | 全ての一致部分を置換（正規表現可） |
| `Reverse()`                                                                | String2 | 文字列を逆順にする         |

---

### 部分文字列操作

| メソッド                                            | 戻り値型     | 説明           |
| ----------------------------------------------- | -------- | ------------ |
| `Slice(indexStart As Long, [indexEnd As Long])` | String2  | 部分文字列（負の値対応） |
| `SubString(starts As Long, [ends As Long])`     | String2  | 開始〜終了の部分文字列  |
| `Split([separator As String], [limit As Long])` | String() | 文字列を区切り文字で分割 |

---

### 前方/後方一致判定

| メソッド                                         | 戻り値型    | 説明          |
| -------------------------------------------- | ------- | ----------- |
| `StartsWith(prefix As String, [position])`   | Boolean | 指定の接頭辞で始まるか |
| `EndsWith(searchString As String, [length])` | Boolean | 指定の接尾辞で終わるか |

---

### トリミング

| メソッド          | 戻り値型    | 説明       |
| ------------- | ------- | -------- |
| `Trim()`      | String2 | 前後の空白を除去 |
| `TrimStart()` | String2 | 前方の空白を除去 |
| `TrimEnd()`   | String2 | 後方の空白を除去 |

---

## 注意事項

* 正規表現を使用する場合は、`VBScript.RegExp` オブジェクトに依存しています。
* 一部メソッド（例：`Repeat`）では負数引数でエラーをスローします。
* 文字列のインデックスは **0ベース** です（VBA標準とは異なる点に注意）。

---

## ライセンス

MIT License
