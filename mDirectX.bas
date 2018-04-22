Attribute VB_Name = "mDirectX"
'---------------------------------------------------------------------------------------
' Module    : mEngine
' DateTime  : 10/30/2004 10:28
' Author    : Shane Mulligan
' Purpose   : Harnesses DirectX8 for use in RS
'---------------------------------------------------------------------------------------

Option Explicit

Public Dx As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8

Public D3DDevice As Direct3DDevice8

Public D3DWindow As D3DPRESENT_PARAMETERS ' Describes the Window
Public DispMode As D3DDISPLAYMODE ' Describes the display mode

Sub InitD3D()

   Set Dx = New DirectX8
   Set D3D = Dx.Direct3DCreate ' Make the D3D object
   Set D3DX = New D3DX8

End Sub

Function Initialise(hWnd As Long) As Boolean

On Local Error GoTo ErrHandler

   D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode  ' retrive the default display
   
   DispMode.Format = D3DFMT_A8R8G8B8
   DispMode.Width = 1024
   DispMode.Height = 768

   D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP ' or maybe D3DSWAPEFFECT_FLIP
   D3DWindow.BackBufferCount = 1 ' 1 backbuffer only
   D3DWindow.BackBufferFormat = DispMode.Format 'What we specified earlier
   D3DWindow.BackBufferHeight = 768
   D3DWindow.BackBufferWidth = 1024
   D3DWindow.Windowed = 1
   D3DWindow.hDeviceWindow = hWnd

   
   ' Create the 3d device
   Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
   Debug.Print "3D DEVICE STARTED"
   
   ' set vertex shaders and such
   D3DDevice.SetVertexShader FVF
   D3DDevice.SetRenderState D3DRS_LIGHTING, False 'Lighting off
   
   ' antialiases textures
   D3DDevice.SetTextureStageState 0, D3DTSS_MAGFILTER, D3DTEXF_LINEAR
   D3DDevice.SetTextureStageState 0, D3DTSS_MINFILTER, D3DTEXF_LINEAR
   
   Initialise = True 'Return success
   
   Exit Function
   
ErrHandler:
   Debug.Print "ERROR IN MODULE ENGINE.INIT"
   Debug.Print Space(5) & Err.Description
   Initialise = False  'Return error
   
End Function

Sub CleanUp()

   On Local Error Resume Next
   Set D3DDevice = Nothing
   Set D3DX = Nothing
   Set D3D = Nothing
   Set Dx = Nothing
   Debug.Print "Engine objects destroyed"

End Sub


   Initialise = False  'Return error
   
End Function

Sub CleanUp()

   On Local Error Resume Next
   Set D3DDevice = Nothing
   Set D3DX = Nothing
   Set D3D = Nothing
   Set Dx = Nothing
   Debug.Print "Engine objects destroyed"

End Sub

