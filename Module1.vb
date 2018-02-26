Imports System.Xml
Imports System.IO
Module Module1
    Public ListX As String
    Sub Main()

        Call A(vbNull)


        'Console.WriteLine(ListX)

        Console.Read()
    End Sub
    Sub B()
        Dim DirInfo As DirectoryInfo
        If Directory.Exists("Mob") Then
            DirInfo = New DirectoryInfo("Mob")
            For Each filex In DirInfo.GetFiles("*.XML")
                '設定抓檔案(*.XML)限定在此資料夾下不包含子資料夾  
                Console.WriteLine(filex.Name) '這就是檔案名稱
                Call A(filex.Name)
            Next filex
        Else
            Console.WriteLine("無此資料夾")
        End If
    End Sub
    Sub A(testfile As String)
        Try
            Dim xDoc As XmlDocument = New XmlDocument
            'xDoc.Load("Mob\" & testfile) '讀取單一檔案
            xDoc.Load("0100100.img.xml") '讀取單一檔案
            Dim xRoot As XmlNode
            Dim xNode As XmlNode
            Dim xIntNodeList As XmlNodeList
            Dim xIntNode As XmlNode
            Dim xElem As XmlElement

            xRoot = CType(xDoc.DocumentElement, XmlNode)
            '選擇imgdir
            xNode = xRoot.SelectSingleNode("imgdir[@name='info']")
            '節點篩選條件
            xIntNodeList = xNode.SelectNodes("int[@name='maxHP' or @name='PADamage' or @name='MADamage']")
            For i = 0 To xIntNodeList.Count - 1
                xIntNode = xIntNodeList.Item(i)
                xElem = CType(xIntNode, XmlElement)

                xElem.SetAttribute("value", Math.Ceiling(xElem.GetAttribute("value") * 0.9))
                Console.WriteLine(xElem.GetAttribute("name") & "：" & xElem.GetAttribute("value"))

                'xElem.Value = 100
            Next
            xDoc.Save("0100100.img.xml")
        Catch ex As Exception
            'Console.WriteLine(testfile & "此檔案沒有節點" & vbNewLine & ex.Message & vbNewLine & ex.StackTrace)
            'MsgBox(testfile & "此檔案沒有節點" & vbNewLine & ex.Message & vbNewLine & ex.StackTrace)
            ListX &= testfile & vbNewLine
        End Try
    End Sub

End Module
