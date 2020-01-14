Attribute VB_Name = "SutYellow"
Option Explicit

' *********************************************************
' SutYellow.dll関連のモジュール
'
' 作成者　：Hideki Isobe
' 履歴　　：2009/05/20　新規作成
'
' 特記事項：
' *********************************************************

#If (DEBUG_MODE = 1) Then
    #If VBA7 And Win64 Then
        Public Declare PtrSafe Function PreLoadIcon Lib ".\..\CPP\Sut\x64\Debug ASM\SutYellow.dll" () As Boolean
        Public Declare PtrSafe Function FreeIcon Lib ".\..\CPP\Sut\x64\Debug ASM\SutYellow.dll" () As Boolean
        Public Declare PtrSafe Function LoadIconAndGetPictureDisp Lib ".\..\CPP\Sut\x64\Debug ASM\SutYellow.dll" (ByVal Id As SutYellowIcons) As IPictureDisp
    #Else
        Public Declare Function PreLoadIcon Lib ".\..\CPP\Sut\Debug ASM\SutYellow.dll" () As Boolean
        Public Declare Function FreeIcon Lib ".\..\CPP\Sut\Debug ASM\SutYellow.dll" () As Boolean
        Public Declare Function LoadIconAndGetPictureDisp Lib ".\..\CPP\Sut\Debug ASM\SutYellow.dll" (ByVal Id As SutYellowIcons) As IPictureDisp
    #End If
#Else
    #If VBA7 And Win64 Then
        Public Declare PtrSafe Function PreLoadIcon Lib "lib\SutYellow.dll" () As Boolean
        Public Declare PtrSafe Function FreeIcon Lib "lib\SutYellow.dll" () As Boolean
        Public Declare PtrSafe Function LoadIconAndGetPictureDisp Lib "lib\SutYellow.dll" (ByVal Id As SutYellowIcons) As IPictureDisp
    #Else
        Public Declare Function PreLoadIcon Lib "lib\SutYellow.dll" () As Boolean
        Public Declare Function FreeIcon Lib "lib\SutYellow.dll" () As Boolean
        Public Declare Function LoadIconAndGetPictureDisp Lib "lib\SutYellow.dll" (ByVal Id As SutYellowIcons) As IPictureDisp
    #End If
#End If

Public Enum SutYellowIcons

    IDB_ICON_ADD = 101
    IDB_ICON_ADD_MASK = 102
    IDB_ICON_ADD_FILE = 103
    IDB_ICON_ADD_FILE_MASK = 104
    IDB_ICON_ADD_FOLDER = 105
    IDB_ICON_ADD_FOLDER_MASK = 106
    IDB_ICON_ADD_FOLDER2 = 107
    IDB_ICON_ADD_FOLDER2_MASK = 108
    IDB_ICON_ALERT = 109
    IDB_ICON_ALERT_MASK = 110
    IDB_ICON_ALERT_MESSAGE = 111
    IDB_ICON_ALERT_MESSAGE_MASK = 112
    IDB_ICON_BOOK = 113
    IDB_ICON_BOOK_MASK = 114
    IDB_ICON_BUTTON_HELP = 115
    IDB_ICON_BUTTON_HELP_MASK = 116
    IDB_ICON_DATABASE = 117
    IDB_ICON_DATABASE_MASK = 118
    IDB_ICON_DATABASE_SETTING = 119
    IDB_ICON_DATABASE_SETTING_MASK = 120
    IDB_ICON_DATABASE_SEARCH = 166
    IDB_ICON_DATABASE_SEARCH_MASK = 167
    IDB_ICON_DELETE = 121
    IDB_ICON_DELETE_MASK = 122
    IDB_ICON_DELETE_DATABASE = 123
    IDB_ICON_DELETE_DATABASE_MASK = 124
    IDB_ICON_DEVIL = 125
    IDB_ICON_DEVIL_MASK = 126
    IDB_ICON_EDIT = 127
    IDB_ICON_EDIT_MASK = 128
    IDB_ICON_REMOVE = 129
    IDB_ICON_REMOVE_MASK = 130
    IDB_ICON_RUN = 131
    IDB_ICON_RUN_MASK = 132
    IDB_ICON_SAVE_AS = 133
    IDB_ICON_SAVE_AS_MASK = 134
    IDB_ICON_SEARCH = 135
    IDB_ICON_SEARCH_MASK = 136
    IDB_ICON_SEARCH_WINDOW = 137
    IDB_ICON_SEARCH_WINDOW_MASK = 138
    IDB_ICON_SETTINGS = 139
    IDB_ICON_SETTINGS_MASK = 140
    IDB_ICON_SMILE = 141
    IDB_ICON_SMILE_MASK = 142
    IDB_ICON_WINDOW_IMPORT = 143
    IDB_ICON_WINDOW_IMPORT_MASK = 144
    IDB_ICON_FLAG_GREEN = 145
    IDB_ICON_FLAG_GREEN_MASK = 146
    IDB_ICON_FLAG_BLUE = 147
    IDB_ICON_FLAG_BLUE_MASK = 148
    IDB_ICON_FLAG_RED = 149
    IDB_ICON_FLAG_RED_MASK = 150
    IDB_ICON_AREA_ADD = 151
    IDB_ICON_AREA_ADD_MASK = 152
    IDB_ICON_AREA_EDIT = 153
    IDB_ICON_AREA_EDIT_MASK = 154
    IDB_ICON_AREA_REMOVE = 155
    IDB_ICON_AREA_REMOVE_MASK = 156
    IDB_ICON_AREA_SEARCH = 157
    IDB_ICON_AREA_SEARCH_MASK = 158
    IDB_ICON_BUG = 159
    IDB_ICON_BUG_MASK = 160
    IDB_ICON_PASTE = 163
    IDB_ICON_PASTE_MASK = 161
    IDB_ICON_FORWARD = 164
    IDB_ICON_FORWARD_MASK = 165

End Enum

