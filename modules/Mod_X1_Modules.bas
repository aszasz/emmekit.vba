Attribute VB_Name = "Mod_X1_Modules"
Private ref As Collection
Private names As Collection

Public Sub update_all_my_basic_modules()
''    If ref Is Nothing Or names Is Nothing Then
        load_names_of_modules
''    End If
    For Each N In names
        update_module CStr(N)
    Next
End Sub

Public Sub export_all_my_basic_modules()
''    If ref Is Nothing Or names Is Nothing Then
        load_names_of_modules
''    End If
    For Each N In names
        export_module CStr(N)
    Next
    Application.VBE.ActiveVBProject.VBComponents("modules").Export "mod\version_control\modules.bas"
End Sub

Public Sub load_names_of_modules()
    Set ref = New Collection: Set names = New Collection
    ref.Add "modules/Mod_A1_Test.bas", "Mod_A1_Test": names.Add "Mod_A1_Test"
    ref.Add "modules/Mod_A2_Dialog.bas", "Mod_A2_Dialog": names.Add "Mod_A2_Dialog"
    ref.Add "modules/Mod_A3_TextTools.bas", "Mod_A3_TextTools": names.Add "Mod_A3_TextTools"
    ref.Add "modules/Mod_A4_Sort.bas", "Mod_A4_Sort": names.Add "Mod_A4_Sort"
    ref.Add "modules/Mod_B1_BasicNetwork.bas", "Mod_B1_BasicNetwork": names.Add "Mod_B1_BasicNetwork"
    ref.Add "modules/Mod_B2_Emme.bas", "Mod_B2_Emme": names.Add "Mod_B2_Emme"
    ref.Add "modules/Mod_C1_Cones.bas", "Mod_C1_Cones": names.Add "Mod_C1_Cones"
    ref.Add "modules/Mod_C2_Line_Mover.bas", "Mod_C2_Line_Mover": names.Add "Mod_C2_Line_Mover"
    ref.Add "modules/Mod_X1_Modules.bas", "Mod_X1_Modules": names.Add "Mod_X1_Modules"
'    ref.Add "Mod_Y_Removed_for_Now.bas", "Mod_Y_Removed_for_Now": names.Add "Mod_Y_Removed_for_Now"
    ref.Add "modules/Mod_Z_VBAConstants.bas", "Mod_Z_VBAConstants": names.Add "Mod_Z_VBAConstants"
    
    On Error Resume Next
    Application.VBE.ActiveVBProject.References.AddFromFile "C:\Windows\System32\vbscript.dll\3" ' Microsoft VBScript Regular Expressions 5.0
    Application.VBE.ActiveVBProject.References.AddFromFile "C:\Windows\system32\scrrun.dll" ' Microsoft Scripting Runtime (for Dictionary support)
    On Error GoTo 0
End Sub

Public Sub unload_module(Name As String)
    On Error Resume Next
        Application.VBE.ActiveVBProject.VBComponents.Remove Application.VBE.ActiveVBProject.VBComponents(Name)
    On Error GoTo 0
End Sub

Public Sub unload_all_my_basic_modules()
    If ref Is Nothing Or names Is Nothing Then
        load_names_of_modules
    End If
    For Each N In names
        unload_module CStr(N)
    Next
End Sub

Public Sub update_module(Name As String)
    unload_module Name
    Application.VBE.ActiveVBProject.VBComponents.import ref(Name)
End Sub
9
Public Sub export_module(Name As String)
    Application.VBE.ActiveVBProject.VBComponents(Name).Export (ref(Name))
End Sub

