Attribute VB_Name = "doc_to_sql_mysql"
Sub doc_to_sql()
    generate_sql ("D:\deployer.sql")
End Sub

'----------------------------------------------
' 支持表格式：
'
' 数据库信息配置表：db_t
'----------------------------------------------
' 字段  |  类型          |   说明
'----------------------------------------------
' id    |  int           |   PK,  AUTO_INCREMENT。数据库ID
'----------------------------------------------
' name  |  varchar(255)  |  数据库名
'----------------------------------------------
'
Sub generate_sql(output)
     Set fsObj = CreateObject("Scripting.FileSystemObject")
     Set file = fsObj.CreateTextFile(output, True)

   For i = 1 To ActiveDocument.Tables.Count
        Set myRegexp = CreateObject("vbscript.regexp")
        
        custom_sql = ActiveDocument.Tables(i).Range.Paragraphs.Last.Next
        tb_name = ActiveDocument.Tables(i).Range.Paragraphs.First.Previous
        If InStr(tb_name, "表：") <= 0 Then
            GoTo nxt
        End If
        
        myRegexp.Pattern = "表：\w+"
        tb_name = Replace(myRegexp.Execute(tb_name)(0), "表：", "")
        
        file.writeline ("drop table if exists " & tb_name & ";")
        file.writeline ("create table " & tb_name & "(")
        
        'myRegexp.Pattern = "[^\r\n]+"
        myRegexp.Pattern = ".+"
        primary_key = "primary key("
        
        For r = 2 To ActiveDocument.Tables(i).Rows.Count
            col1 = myRegexp.Execute(ActiveDocument.Tables(i).Cell(r, 1).Range.Text)(0)
            col2 = myRegexp.Execute(ActiveDocument.Tables(i).Cell(r, 2).Range.Text)(0)
            col3 = myRegexp.Execute(ActiveDocument.Tables(i).Cell(r, 3).Range.Text)(0)
            
            col1 = Replace(Replace(col1, Chr(7), ""), Chr(13), "")
            col2 = Replace(Replace(col2, Chr(7), ""), Chr(13), "")
            
            '判断注释部分是否有多个换行符，若有则保留换行符，否则清掉换行符
            arr = Split(col3, Chr(13))
            If UBound(arr) > 1 Then
                col3 = Replace(col3, Chr(7), "")
            Else
                col3 = Replace(Replace(col3, Chr(7), ""), Chr(13), "")
            End If
                   
            '带下划线表示主键
            If ActiveDocument.Tables(i).Cell(r, 1).Range.Font.Underline Then
                '斜字体表示自增
                If ActiveDocument.Tables(i).Cell(r, 1).Range.Font.Italic Then
                    file.writeline (col1 & " " & col2 & " not null auto_increment comment '" & col3 & "',")
                Else
                    file.writeline (col1 & " " & col2 & " not null comment '" & col3 & "',")
                End If
                
                If r = 2 Then
                    primary_key = primary_key & col1
                Else
                    primary_key = primary_key & "," & col1
                End If
            Else
                file.writeline (col1 & " " & col2 & " comment '" & col3 & "',")
            End If
        Next
        
        file.writeline (primary_key & ")")
        file.writeline (") ENGINE=InnoDB DEFAULT CHARSET=utf8;")
        file.writeline ("")
        
        If InStr(custom_sql, "<SQL>") > 0 Then
            tp = custom_sql.Paragraphs.First.Next
            While InStr(tp, "</SQL>") <= 0
                Debug.Print "Debug:" & tp
                file.writeline (tp)
                   
                tp = tp.Paragraphs.First.Next
            Wend
        End If
    
nxt:
   Next

End Sub
