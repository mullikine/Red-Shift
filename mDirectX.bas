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

TEXMODE = 140
    D3DRS_COLORVERTEX = 141
    D3DRS_LOCALVIEWER = 142
    D3DRS_NORMALIZENORMALS = 143
    D3DRS_DIFFUSEMATERIALSOURCE = 145
    D3DRS_SPECULARMATERIALSOURCE = 146
    D3DRS_AMBIENTMATERIALSOURCE = 147
    D3DRS_EMISSIVEMATERIALSOURCE = 148
    D3DRS_VERTEXBLEND = 151
    D3DRS_CLIPPLANEENABLE = 152
    D3DRS_POINTSIZE = 154
    D3DRS_POINTSIZE_MIN = 155
    D3DRS_POINTSPRITEENABLE = 156
    D3DRS_POINTSCALEENABLE = 157
    D3DRS_POINTSCALE_A = 158
    D3DRS_POINTSCALE_B = 159
    D3DRS_POINTSCALE_C = 160
    D3DRS_MULTISAMPLEANTIALIAS = 161
    D3DRS_MULTISAMPLEMASK = 162
    D3DRS_PATCHEDGESTYLE = 163
    D3DRS_DEBUGMONITORTOKEN = 165
    D3DRS_POINTSIZE_MAX = 166
    D3DRS_INDEXEDVERTEXBLENDENABLE = 167
    D3DRS_COLORWRITEENABLE = 168
    D3DRS_TWEENFACTOR = 170
    D3DRS_BLENDOP = 171
    D3DRS_POSITIONDEGREE = 172
    D3DRS_NORMALDEGREE = 173
    D3DRS_SCISSORTESTENABLE = 174
    D3DRS_SLOPESCALEDEPTHBIAS = 175
    D3DRS_ANTIALIASEDLINEENABLE = 176
    D3DRS_MINTESSELLATIONLEVEL = 178
    D3DRS_MAXTESSELLATIONLEVEL = 179
    D3DRS_ADAPTIVETESS_X = 180
    D3DRS_ADAPTIVETESS_Y = 181
    D3DRS_ADAPTIVETESS_Z = 182
    D3DRS_ADAPTIVETESS_W = 183
    D3DRS_ENABLEADAPTIVETESSELLATION = 184
    D3DRS_TWOSIDEDSTENCILMODE = 185
    D3DRS_CCW_STENCILFAIL = 186
    D3DRS_CCW_STENCILZFAIL = 187
    D3DRS_CCW_STENCILPASS = 188
    D3DRS_CCW_STENCILFUNC = 189
    D3DRS_COLORWRITEENABLE1 = 190
    D3DRS_COLORWRITEENABLE2 = 191
    D3DRS_COLORWRITEENABLE3 = 192
    D3DRS_BLENDFACTOR = 193
    D3DRS_SRGBWRITEENABLE = 194
    D3DRS_DEPTHBIAS = 195
    D3DRS_WRAP8 = 198
    D3DRS_WRAP9 = 199
    D3DRS_WRAP10 = 200
    D3DRS_WRAP11 = 201
    D3DRS_WRAP12 = 202
    D3DRS_WRAP13 = 203
    D3DRS_WRAP14 = 204
    D3DRS_WRAP15 = 205
    D3DRS_SEPARATEALPHABLENDENABLE = 206
    D3DRS_SRCBLENDALPHA = 207
    D3DRS_DESTBLENDALPHA = 208
    D3DRS_BLENDOPALPHA = 209
    D3DRS_FORCE_DWORD = &H7FFFFFFF
End Enum

Sub InitD3D()

   Set Dx = New DirectX8
   Set D3D = Dx.Direct3DCreate ' Make the D3D object
   Set D3DX = New D3DX8

End Sub

Function Initialise(hWnd As Long) As Boolean

On Local Error GoTo ErrHandler

   D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode  ' retrive the default display
   
   DispMode.Format = COLOR_DEPTH_32_BIT
   DispMode.Width = Screen.Width / Screen.TwipsPerPixelX
   DispMode.Height = Screen.Height / Screen.TwipsPerPixelY

   D3DWindow.SwapEffect = D3DSWAPEFFECT_FLIP ''D3DSWAPEFFECT_COPY_VSYNC 'refresh when monitor does , or D3DSWAPEFFECT_FLIP
   D3DWindow.BackBufferCount = 1 ' 1 backbuffer only
   D3DWindow.BackBufferFormat = DispMode.Format 'What we specified earlier
   D3DWindow.BackBufferWidth = DispMode.Width
   D3DWindow.BackBufferHeight = DispMode.Height
   D3DWindow.Windowed = bFullscreen
   D3DWindow.hDeviceWindow = hWnd
   
   ' Create the 3d device
   Set D3DDevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)
   
   Debug.Print "3D DEVICE STARTED"
   
   ' Let both clockwise and anti-clockwise triangles be drawn
   D3DDevice.SetRenderState D3DRS_CULLMODE, D3DCULL_NONE '
   
   ' set vertex shaders and such
   D3DDevice.SetVertexShader FVF
   D3DDevice.SetRenderState D3DRS_LIGHTING, False 'Lighting off
   
   ' Turn off the zbuffer
   D3DDevice.SetRenderState D3DRS_ZENABLE, True
   
   ' uses material alpha
   ''D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_SELECTARG1
   ''D3DDevice.SetTextureStageState 0, D3DTSS_ALPHAARG1, D3DTA_DIFFUSE
   
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

'Function PostInitialize(ByVal WindowWidth As Single, ByVal WindowHeight As Single)
'
'   Dim Ortho2D As D3DXMATRIX
'   D3DXMATRIX Identity;
'
'   D3DXMatrixOrthoLH(&Ortho2D, WindowWidth, WindowHeight, 0.0f, 1.0f);
'   D3DXMatrixIdentity(&Identity);
'
'   g_pd3dDevice->SetTransform(D3DTS_PROJECTION, &Ortho2D);
'   g_pd3dDevice->SetTransform(D3DTS_WORLD, &Identity);
'   g_pd3dDevice->SetTransform(D3DTS_VIEW, &Identity);
'
'End Function

Sub CleanUp()

   On Local Error Resume Next
   Set D3DDevice = Nothing
   Set D3DX = Nothing
   Set D3D = Nothing
   Set Dx = Nothing
   Debug.Print "Engine objects destroyed"

End Sub

