script type
ftplugin
 
description
Support Automatic functions like Emacs for Verilog HDL
Feature list:
1) Auto Argument (the same as Emacs)
2) Auto Instance (power than Emacs)
3) Auto Define Signals
4) Auto unit delay "<=" to "<= #`FFD"
5) Auto always block
6) Auto header
7) Auto comment
My blog: http://blog.sina.com.cn/arrowroothover
My E-mail: arrowroothover@hotmail.com
 
install details
put the automatic.vim in .vim/ftplugin/vlog/
befor.v: a Verilog example before run auto functions
after.v: a Verilog example after run auto functions

Private Sub Application_Quit()
  If TimerID <> 0 Then Call DeactivateTimer 'Turn off timer upon quitting **VERY IMPORTANT**
End Sub

Private Sub Application_Startup()
    result = MsgBox("Activate the Mail Filter?" & Chr(10) & "*If you want to upgrade the Rule, please select NO", vbYesNo)
    If result = vbYes Then
        Call ActivateTimer(5) 'Set timer to go off every 5 seconds
    Else
        result = MsgBox("If you want to activate Mail Filter, you should restart the Outlook!", vbOKOnly, "Filtering Mail by VBA")
    End If
End Sub
