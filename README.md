# CatiaPythonMastery
CATIA V5とPythonを使って、図面ビューから投影面を作成する方法を学びました。

参考にした動画は[こちら](https://youtu.be/ya35YTEf580?si=2kOOUqb_gIMs5f3C)です。

#### 使用したコードとその詳細説明

```python
import win32com.client  # pywin32パッケージが必要です
from sys import exit

# CATIAを取得
CATIA = win32com.client.Dispatch("catia.Application")

# アクティブなドキュメント名を取得
activeDocName = CATIA.activedocument.name

# 現在のドキュメントが.CATDrawingか確認
if ".CATDrawing" not in activeDocName:
  print("現在開いているドキュメントは図面ではありません")
  exit()

# 選択を取得
oSel = CATIA.activedocument.selection

# 選択されたアイテムがあるか確認
if oSel.count > 0:
  for i in range(oSel.count):
     selection = oSel.item(i + 1)
     typeDoc = selection.type
     if typeDoc == "DrawingView":
       view = oSel.item(i + 1).value
       parentLink = view.generativelinks.firstlink().parent.name
       partList.append(parentLink)
       
       X1, Y1, Z1, X2, Y2, Z2 = view.GenerativeBehavior.GetProjectionPlane()
       
       dataView[view.name] = (X1, Y1, Z1, X2, Y2, Z2)

# CatPartを取得
thispart = CATIA.documents.item(partList[0]).partList

# ジオセットを作成
geoSet = thispart.HybridBodies.add()
geoSet.name = "Export Plane from 2D"

hsb = thispart.HybridShapeFactory

# 最初のポイントを作成
pointMain = hsb.AddNewPointCoord(0, 0, 0)
pointMain.name = "Origin"
ref = thispart.CreateReferenceFromObject(pointMain)
geoSet.AppendHybridShape(pointMain)

thispart.update()

for item, value in dataView.items():
  dirByCoord1 = hsb.AddNewDirectionByCoord(value[0], value[1], value[2])
  dirByCoord2 = hsb.AddNewDirectionByCoord(value[3], value[4], value[5])

  LinepointDir1 = hsb.AddNewLinePtDir(ref, dirByCoord1, 0, 0, 35, False)
  LinepointDir1.name = item + "X"
  geoSet.AppendHybridShape(LinepointDir1)

  LinepointDir2 = hsb.AddNewLinePtDir(ref, dirByCoord2, 0, 0, 35, False)
  LinepointDir2.name = item + "Y"
  geoSet.AppendHybridShape(LinepointDir2)

  planeLine = hsb.AddNewPlane2Lines(LinepointDir1, LinepointDir2)
  planeLine.name = item
  geoSet.AppendHybridShape(planeLine)
  thispart.update()
```

### コードの詳細説明

1. **ライブラリのインポート**
   ```python
   import win32com.client  # pywin32パッケージが必要です
   from sys import exit
   ```
   - `win32com.client`は、Windows COMオブジェクトを操作するためのライブラリです。`pywin32`をインストールする必要があります。

2. **CATIAアプリケーションの取得**
   ```python
   CATIA = win32com.client.Dispatch("catia.Application")
   ```
   - CATIAアプリケーションのインスタンスを取得します。

3. **アクティブなドキュメントの確認**
   ```python
   activeDocName = CATIA.activedocument.name
   if ".CATDrawing" not in activeDocName:
     print("現在開いているドキュメントは図面ではありません")
     exit()
   ```
   - 現在アクティブなドキュメントがCATIAの図面ファイルであるか確認します。

4. **選択されたアイテムの処理**
   ```python
   oSel = CATIA.activedocument.selection
   if oSel.count > 0:
     for i in range(oSel.count):
        selection = oSel.item(i + 1)
        typeDoc = selection.type
        if typeDoc == "DrawingView":
          view = oSel.item(i + 1).value
          parentLink = view.generativelinks.firstlink().parent.name
          partList.append(parentLink)
          
          X1, Y1, Z1, X2, Y2, Z2 = view.GenerativeBehavior.GetProjectionPlane()
          
          dataView[view.name] = (X1, Y1, Z1, X2, Y2, Z2)
   ```
   - 選択されたアイテムがある場合、その中から`DrawingView`タイプのものを特定し、ビューの投影面の座標を取得します。

5. **CatPartの取得とジオセットの作成**
   ```python
   thispart = CATIA.documents.item(partList[0]).partList
   geoSet = thispart.HybridBodies.add()
   geoSet.name = "Export Plane from 2D"
   ```
   - 取得したビューのリンク元であるCatPartを取得し、ジオセットを作成します。

6. **ポイント、方向、ライン、平面の作成**
   ```python
   pointMain = hsb.AddNewPointCoord(0, 0, 0)
   pointMain.name = "Origin"
   ref = thispart.CreateReferenceFromObject(pointMain)
   geoSet.AppendHybridShape(pointMain)

   for item, value in dataView.items():
     dirByCoord1 = hsb.AddNewDirectionByCoord(value[0], value[1], value[2])
     dirByCoord2 = hsb.AddNewDirectionByCoord(value[3], value[4], value[5])

     LinepointDir1 = hsb.AddNewLinePtDir(ref, dirByCoord1, 0, 0, 35, False)
     LinepointDir1.name = item + "X"
     geoSet.AppendHybridShape(LinepointDir1)

     LinepointDir2 = hsb.AddNewLinePtDir(ref, dirByCoord2, 0, 0, 35, False)
     LinepointDir2.name = item + "Y"
     geoSet.AppendHybridShape(LinepointDir2)

     planeLine = hsb.AddNewPlane2Lines(LinepointDir1, LinepointDir2)
     planeLine.name = item
     geoSet.AppendHybridShape(planeLine)
     thispart.update()
   ```
   - 原点となるポイントを作成し、それに基づいて方向、ライン、そして平面を作成します。取得したデータを基に、それぞれの方向とラインを設定し、最後に平面を作成します。

このようにして、CATIA V5とPythonを組み合わせて、図面ビューから投影面を作成するプロセスを理解しました。
