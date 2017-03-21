main

Sub main()
    Dim target_name
    target_name  = InputBox("doujinshi search")

    If target_name = "" then
        Exit Sub
    End If

    search_toranoana target_name
    search_melonbooks target_name
    search_comic_zin target_name

End Sub

Sub search_toranoana(name)

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True

    ie.Navigate "http://www.toranoana.jp/mailorder/"
    waitIE ie

    ie.Document.getElementsByName("search")(0).Value = name
    WScript.Sleep 100

    Set inputs = ie.Document.getElementsByTagName("input")
    For i = 0 To inputs.Length - 1
        If inputs(i).type = "submit" then
            inputs(i).Click
            Exit For
        End If
    Next
    waitIE ie

End Sub

Sub search_melonbooks(name)

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True

    ie.Navigate "https://www.melonbooks.co.jp/?adult_auth=1"
    waitIE ie

    On Error Resume Next
    ie.Document.getElementsByClassName("f_left yes")(0).Click
    waitIE ie
    On Error Goto 0

    ie.Document.getElementsByClassName("input rich")(0).Value = name
    WScript.Sleep 100

    ie.Document.getElementsByClassName("submit")(0).Click
    waitIE ie

    ie.Document.getElementsByName("is_end_of_sale[]")(0).Checked = True
    waitIE ie

    ie.Document.getElementsByClassName("submit submit_small")(0).Click
    waitIE ie
End Sub

Sub search_comic_zin(name)

    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True

    ie.Navigate "http://shop.comiczin.jp/products/list.php"
    waitIE ie

    ie.Document.getElementById("txt_search_word").Value = name
    WScript.Sleep 100

    ie.Document.getElementById("btn_search").Click
    waitIE ie

End Sub

Sub waitIE(ie)

    Do While ie.Busy = True Or ie.readystate <> 4
        WScript.Sleep 100
    Loop

End Sub
