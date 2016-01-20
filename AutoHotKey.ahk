SendMode, Input
#InstallKeybdHook
#singleinstance force
#Persistent   ; without this line the script will exit
KeyHistory

; CUSTOM MENU -------------------------------------------------------------------------------------------------------
Menu, MyMenu, Add, CB Accept, CBAccAction
Menu, MyMenu, Add, Item2, Item2Action
Menu, MyMenu, Add  ; Add a separator line.
Menu, MyMenu, Add, Item3, Item3Action
Menu, MyMenu, Add, Item4, Item4Action
; Create another menu destined to become a submenu of the above menu.
Menu, Submenu1, Add, Item1, SubMenu1Action
Menu, Submenu1, Add, Item2, SubMenu2Action
Menu, MyMenu, Add, Test Submenu, :Submenu1
Menu, MyMenu, Add  ; Add a separator line.
Menu, MyMenu, Add, MenuClose, MenuCloseAction

return  ; End of script's auto-execute section.

CBAccAction:
accvar = balance, not representing, added to CB and logging, deactivating to change status
SendInput %accvar% {SHIFT DOWN}{TAB 2}{SHIFT UP}{RIGHT}{TAB} 
Sleep 10
SendInput ch 
Sleep 10
SendInput {TAB 2}{LEFT}{TAB 2}	
return

Item2Action:
Send Item2 Test
return

Item3Action:
Send Item3 Test
return

Item4Action:
Send Item4 Test
return

SubMenu1Action:
Send Sub1Test
return

SubMenu2Action:
Send Sub2Test
return

MenuCloseAction:
return

XButton1::RButton
;XButton1::
Menu, MyMenu, Show
return

; CHECK WHICH PLATFORM COPIED CASE # IS FROM ------------------------------------------------------------------------
#F1::
lastClip :=ClipBoard
StringLen, length, lastClip
if length = 14
MsgBox, This Case # is from ClientLine!
if length = 15
MsgBox, This Case # is from PayPal!
if length = 16
MsgBox, This Case # is from Adyen!
if length = 7
MsgBox, This Case # is from Amex!
return

; DOUBLE CLICK TO COPY-----------------------------------------------------------------------------------------------
~LButton::
If (A_PriorHotKey = A_ThisHotKey and A_TimeSincePriorHotkey < 400)
{
Send ^c ; or your double-right-click action here
}
Return

; MIDDLE CLICK PASTE PLAIN TEXT--------------------------------------------------------------------------------------
~MButton::
Clip0 = %ClipBoardAll%
   ClipBoard = %ClipBoard%       ; Convert to text
   Send ^v                       ; For best compatibility: SendPlay
   Sleep 50                      ; Don't change clipboard while it is pasted! (Sleep > 0)
   ClipBoard = %Clip0%           ; Restore original ClipBoard
   VarSetCapacity(Clip0, 0)      ; Free memory
Return

; RIGHT CLICK WITH LEFT MOUSE HELD DOWN TO COPY----------------------------------------------------------------------
RButton::
     if GetKeyState("LButton", "P")
     SendInput ^c
	 		

; RISK HOTKEYS-------------------------------------------------------------------------------------------------------
::2c::2 different cards with AVS Y and same billing,  
::bi::biz ISP/ORG,
::ci::custom image: good
::fb::Facebook match 
::fs::future send, 
::ges::good email syntax, 
::gm::good message
::om::okay message
::ha::high amount for brand,
::hd::high distance, 
::hr::high-risk brand,
::la::low amount for brand,
::ld::low distance, 
::lm::low MF, 
::lr::lower-risk brand,
::mc::multiple different cards with AVS Y and same billing, 
::mo::mobile order,
::na::normal amount for brand,
::nem::name/email mismatch,
::obo::{#}OB call to number on order, 
::obw::{#}OB call to number on WP page, 
::pp::paypal transaction,
::sa::specific amount,
::si::specific item mentioned in message, 
::ssp::self-send paypal,
::wpa::White Pages address match, 
::wpp::White Pages phone match, 
::mn::multiple names/cards/addresses,
::pmf::proxy, high MF,
::noe::No other EGC purchases planned at this time
::oks::okay syntax,
::psw::paid sender with matching ISP,
::bsw::biz sender with matching ISP,
::gms::Google matches sender email,
::gmr::Google matches recip email,
::bms::biz Google matches sender email,
::bmr::biz Google matches recip email,
::res::{#}RESELLER
::dnf::did not feel the need for a full verification
::pv::previously verified
::okp::okay pattern
::okx::okay history
::mfo::multiple forms of payment
::ccc::Please feel free to contact our customer care department at 855-741-1209, Option 2, if you have any other questions or concerns.
::refacc::Issued a manual refund for the accessory charge on this transaction as requested by Customer Care.
::refship::Issued a manual refund for the shipping charge on this transaction as requested by Customer Care.


; MISC. PAYMENTS-----------------------------------------------------------------------------------------------------
::uprr::
uprvar = UPR Report (Unprocessed Returns Report: A credit related to this transaction may have failed to transmit) {TAB}
SendInput %uprvar%
Return

:*:fdr::
fdvar = FD Report (Failed Deactivations Report: An eGC on this transaction may have failed to deactivate) {TAB}
SendInput %fdvar%
Return

::failcred::
failcredvar = Automatic credit that typically occurs when an eGC is returned was not present. Issued a manual credit.
SendInput %failcredvar%
Return

:*:sae::Sent an email to the brand

:*:fdb::
fdbvar = Final deactivation balance(s) recorded
SendInput %fdbvar%
return

:*:priorcred::
priorcredvar = This cardholder has already been refunded for this transaction. Proof via payment software screenshot is attached.
SendInput %priorcredvar%
return

:*:ppp::
paypalvar = paypal{+}@cashstar.com
SendInput ^a
SendInput %paypalvar%
SendInput {Left 13}
return

; CHARGEBACK CLOSING NOTES-------------------------------------------------------------------------------------------
::acc::
accvar = balance, not representing, added to CB and logging, deactivating to change status
SendInput %accvar% {SHIFT DOWN}{TAB 2}{SHIFT UP}{RIGHT}{TAB} 
Sleep 10
SendInput ch 
Sleep 10
SendInput {TAB 2}{LEFT}{TAB 2}
return

:*:rcred::
rcredvar = balance, representing (prior credit), added to CB and logging
SendInput %rcredvar% {TAB}{TAB}
return

:*:rfra::
rfravar = balance, representing (do not believe this transaction is fraudulent), added to CB and logging
SendInput %rfravar% {TAB}{TAB}
return

:*:rrec::
rrecvar = balance, representing (merchandise received and/or redeemed), added to CB and logging
SendInput %rrecvar% {TAB}{TAB}
return

; CHARGEBACK OBJECT HOTKEY-----------------------------------------------------------------------------------------
XButton2::
today = %a_now%
today += -1, days
FormatTime, today, %today%, MM/dd/yyyy 

ClipBoard = %ClipBoard%
lastClip :=ClipBoard

StringLen, length, lastClip
if length < 7
return
if length > 16
return
if length = 15 ;If the case # is from PayPal
{
SendInput {CTRL DOWN}{SHIFT DOWN}{LEFT 3}{SHIFT UP}{CTRL UP}%today%{TAB}x{TAB 2}
Sleep 10
SendInPut x{TAB}
Sleep 10
SendInput ^v {Tab}x
Sleep 10
SendInput {Tab 3}f
Sleep 10
SendInput {Shift Down}{Tab 8}{Shift Up}{RIGHT}{LEFT 5}{Shift Down}{LEFT 2}{SHIFT UP}
}
else ;If the case # is NOT from PayPal
{
SendInput {CTRL DOWN}{SHIFT DOWN}{LEFT 3}{SHIFT UP}{CTRL UP}%today%{TAB 3}
Sleep 10
SendInPut x{TAB}
Sleep 10
SendInput ^v {Tab}x
Sleep 10
SendInput {Tab 3}f
Sleep 10
SendInput {Shift Down}{Tab 8}{Shift Up}{RIGHT}{LEFT 5}{Shift Down}{LEFT 2}{SHIFT UP}
}
return

:*:ddd::
today = %a_now%
today += -1, days
FormatTime, today, %today%, MM/dd/yyyy 

ClipBoard = %ClipBoard%
lastClip :=ClipBoard

StringLen, length, lastClip
if length < 7
return
if length > 16
return
if length = 15 ;If the case # is from PayPal
{
SendInput {CTRL DOWN}{SHIFT DOWN}{LEFT 3}{SHIFT UP}{CTRL UP}%today%{TAB}x{TAB 2}
Sleep 10
SendInPut x{TAB}
Sleep 10
SendInput ^v {Tab}x
Sleep 10
SendInput {Tab 3}f
Sleep 10
SendInput {Shift Down}{Tab 8}{Shift Up}{RIGHT}{LEFT 5}{Shift Down}{LEFT 2}{SHIFT UP}
}
else ;If the case # is NOT from PayPal
{
SendInput {CTRL DOWN}{SHIFT DOWN}{LEFT 3}{SHIFT UP}{CTRL UP}%today%{TAB 3}
Sleep 10
SendInPut x{TAB}
Sleep 10
SendInput ^v {Tab}x
Sleep 10
SendInput {Tab 3}f
Sleep 10
SendInput {Shift Down}{Tab 8}{Shift Up}{RIGHT}{LEFT 5}{Shift Down}{LEFT 2}{SHIFT UP}
}
return

; RANDOM STUFF ------------------------------------------------------------------------------------------------------
SetCapsLockState, AlwaysOff 
SetScrollLockState, AlwaysOff 
Capslock::Shift
#SPACE::  Winset, Alwaysontop, , A
