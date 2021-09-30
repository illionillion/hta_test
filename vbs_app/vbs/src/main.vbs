sub aa()
    msgbox "a"
end sub

Sub Window_OnLoad()
    Window.ResizeTo 340,200
End Sub

dim num
num=0

sub plus()
    num=num+1
    'msgbox num
    document.getElementById("num_area").value=num
end sub

sub minus()

    if num<>0 then
        num=num-1
        'msgbox num
        document.getElementById("num_area").value=num
    end if
end sub

sub reset()
    num=0
    'msgbox num
    document.getElementById("num_area").value=num
end sub