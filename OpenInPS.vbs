Set z=CreateObject("Illustrator.Application")
Set pz=CreateObject("Photoshop.Application")
For Each it In z.ActiveDocument.Selection
If TypeName(it)="PlacedItem" Then pz.Open it.File
Next