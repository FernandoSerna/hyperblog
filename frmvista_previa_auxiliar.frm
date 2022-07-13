VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form frmvista_previa_auxiliar 
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   16305
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   16305
   StartUpPosition =   3  'Windows Default
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer cr2 
      Height          =   8955
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   18840
      _cx             =   33232
      _cy             =   15796
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   2058
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "frmvista_previa_auxiliar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
   Unload Me
End Sub

Private Sub Form_Load()
   'Unload Me
End Sub
