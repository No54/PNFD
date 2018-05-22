Imports System.IO
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Interop

Public Class PNFD

    <CommandMethod("PNFD", CommandFlags.Session)>
    Public Sub PNFDmain()

        Try

            '首先获得要操作的文件夹
            Dim OpPath As String = InputBox("请给出要处理的图纸目录，如（C:\DWG）。" & vbLf & "本插件为测试版本，请注意备份文件。")

            If OpPath = "" Then
                MsgBox("路径不能为空！")
                Exit Sub
            End If

            Dim OutPath As String = OpPath & "\PNFDout"

            If Dir(OutPath, vbDirectory) <> "" Then
                Dim attr1 As FileAttributes = File.GetAttributes(OutPath)
                If (attr1 = FileAttributes.Directory) Then
                    Directory.Delete(OutPath, True)
                Else
                    File.Delete(OutPath)
                End If
            End If

            Directory.CreateDirectory(OutPath & "\")

            Dim tempacadapp As AcadApplication = GetObject(, "Autocad.Application") '检查AutoCAD是否已经打开

            Dim tempdoc As Document = Application.DocumentManager.MdiActiveDocument

            Dim IntF As Integer = 0                                      '用来定义文件总个数

            Dim Fdir As DirectoryInfo = New IO.DirectoryInfo(OpPath)      '由输入的文本路径，获得实际的路径
            For Each fi In Fdir.GetFiles()                                '对于路径内的每个PN文件进行操作

                If Mid(fi.FullName, Len(fi.FullName) - 2, 3) = "dwg" Or Mid(fi.FullName, Len(fi.FullName) - 2, 3) = "DWG" Then
                    IntF = IntF + 1

                    Dim CLineJ As Integer = 0          '找到路径文本里最后一个“\”
                    For j = 1 To Len(fi.FullName)
                        If Mid(fi.FullName, j, 1) = "\" Then
                            CLineJ = j
                        End If
                    Next

                    Dim DocPath As String = Mid(fi.FullName, 1, CLineJ - 1) & "\"                              '前边的是路径
                    Dim FileName As String = Mid(fi.FullName, CLineJ + 1, Len(fi.FullName) - CLineJ - 4)   '后边的是文件名,不包括后缀名

                    Dim acDocMgr As DocumentCollection = Application.DocumentManager

                    Dim doc As Document = DocumentCollectionExtension.Open(acDocMgr, fi.FullName, False)   '打开文件
                    Application.DocumentManager.MdiActiveDocument = doc

                    Dim Adoc As AcadDocument = DocumentExtension.GetAcadDocument(doc)

                    Dim acCurDb As Database = doc.Database   '当前文档数据库

                    Dim m_DocumentLock As DocumentLock = doc.LockDocument()
                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()     '开启一个事务

                        Dim lt As LayerTable = TryCast(acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForWrite), LayerTable)

                        For Each ltid As ObjectId In lt
                            Dim ltr As LayerTableRecord = TryCast(acTrans.GetObject(ltid, OpenMode.ForWrite), LayerTableRecord)
                            If ltr.Name.ToString = "视口" Then
                                ltr.IsOff = True
                            End If
                        Next

                        acTrans.Commit()

                    End Using

                    Using ccTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                        Dim ed As Editor = doc.Editor

                        Dim p1 As Point2d = New Point2d(-55.4355, 1363.28)
                        Dim p2 As Point2d = New Point2d(785.5645, 1957.28)

                        Dim hh As Double = p2.Y + p1.Y
                        Dim ww As Double = p2.X + p1.X

                        Dim pC As Point2d = New Point2d(ww / 2, hh / 2)

                        Dim acView As ViewTableRecord = ed.GetCurrentView()
                        acView.CenterPoint = pC
                        acView.Width = ww / 4
                        acView.Height = hh / 4
                        ed.SetCurrentView(acView)

                        ccTrans.Commit()
                    End Using

                    Dim stroutpath As String = OutPath & "\" & FileName & ".dwg"

                    m_DocumentLock.Dispose()

                    Adoc.SaveAs(stroutpath)

                    Adoc.Close(False)           '关闭报错

                End If
            Next

            MsgBox("运行完成!")

        Catch ex As System.Exception
            MsgBox(ex.ToString)
            Exit Sub

        End Try

    End Sub

End Class
