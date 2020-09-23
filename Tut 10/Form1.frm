VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "NemoX tutorial #10 - Adding Model made with Milkshape3D to scene"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'========================================
' -NemoX tutorial #10 - Adding Model made with Milkshape3D to scene
'========================================
'
'
'Welcome to the NemoX engine Tutorial Series
'
'This is the Tutorial #10 - Adding Model made with Milkshape3D to scene
'for questions and remark go to
'the Engine Website:http://perso.wanadoo.fr/malakoff/index.htm
'
'Or mail me at Johna_pop@yahoo.fr
'
'If you are interresting of making demos,sample,tuts for NemoX engine
'or you want to help me mail Me Johna_pop@yahoo.fr
'
'======================================================
'MAKE SURE YOU HAVE DOWNLOAD THE ENGINE FILES AND HAVE
'SEEN THE TUT #1 at http://perso.wanadoo.fr/malakoff/index.htm
'project section
'=======================================================
'
'
'KEY Arrow to Move left,right,Forward,BackWard maintain Ctrl key to move quickly
'KEY Numpad 8/2  turn UP/DOWN
'KEY Numpad +/- to move camera UP/DOWN
'
'
'====================IMPORTANT=========================================
'   IF U ALREADY HAVE SOME TUTORIALS AND HAVE DOWNLOADED THIS
'   CURRENT ONE, YOU NEED TO UPDATE THE MAIN ENGINE CORE DLL
'   DOWNLOAD IT,DELETE REFERENCE TO OLD PROJECTS AND RE-ADD REFERENCE TO OLD TUTS
'
'   YOU CAN USE SETUP TO DESINSTALL OLD ENGINE VERSION
'   AFTER YOU HAVE DONE THAT INSTALL THE NEW PACKAGE
'
'=================================================================



'=======================================================
'  Tutorial Targets
'
'
'1) --------->how to load Milkshape3D Model file to NemoX Scene
'    go at sub Load_Meshes
'2) --------->how to rescale move,rotate object to Add in the scene
'    go at sub Load_Meshes
'3) --------->how to render the mesh in the scene
'    go at sub Sub gameLoop()
'4) --------->handle with collision detection & Response
'    go at sub Sub CheckCollision()
'5)----------> last word
'  any problems,johna_pop@yahoo.fr
'======================================









'===========Prerequisities


'I assumes you have followed the old step by step tutorials
'if not check at TUT #1 at http://perso.wanadoo.fr/malakoff/index.htm
'

'The Main One is The NemoX renderer
Dim Nemo As NemoX


'we're gonna use a Nemo class for rendering a mesh or polygons
'we use cNemo_Mesh class

Dim WithEvents Mesh As cNemo_Mesh
Attribute Mesh.VB_VarHelpID = -1




'we will need a quick acces to very important and useful functions
Dim Tool As cNemo_Tools




'========NEW OBJECTS=======

'here is the class for our Milkshape3D objects
Dim Milk1 As cNemo_MSD3mesh
Dim Milk2 As cNemo_MSD3mesh
Dim Milk3 As cNemo_MSD3mesh


'New in NemoX 1.72
'this will handle camera moves
Dim KAMERA As cNemo_Camera
'this will handle for keyboard input state
Dim KEY As cNemo_Input


'some constant for player speed
'Check at Sub GetKey()
Private Const RotationSpeed = 6
Private Const MoveSpeed = 2




'Entry point for our project
Private Sub Form_Load()

   Me.Show
  Me.Refresh

 
 


  'call the initializer sub
  Call InitEngine
  
  
  
  'load  MS3D meshes
  Call Load_Meshes
  
    'make geometry
  Call BuilGeometry
  
  
  
  'call the main game Loop
   Call gameLoop

End Sub





'we build our mesh here

Sub BuilGeometry()
      
        
        'adding a plane surface for a simple floor$
        
        'first off very important we pass a texture to the meshbuilder
        Mesh.Add_Texture (App.Path + "\ground.jpg")          '0
        
    
        'Add flooor surface here
        Mesh.Add_WallFloor Tool.Vector(-5000, -1, -5000), Tool.Vector(5000, -1, 5000), 10, 10, 0
        
       
        'Just add some details at scene
       
      
            
      'Feel free to add more geometry details
       For i = 1 To 5
          Mesh.Add_Cilynder Tool.Vector(450 - i * 450 - 50, 10, -815 + 1500), 50, 490, 8, 0
          
          Mesh.Add_Cilynder Tool.Vector(450 - i * 450 - 50, 10, -1035 + 1500), 50, 490, 8, 0
         
      Next i
      
      '========IMPORTANT========
        'then we build our mesh
        Mesh.BuilMesh
        
        
        'to save mesh
       'Mesh.SaveMesh App.Path + "\Level.Nmsh"
       
  
End Sub




Sub Load_Meshes()
  
  'load  mesh #1
  Milk1.LoadMSD3 App.Path + "\Rock\rock.ms3d"
  Milk1.Set_Scale 7, 5, 7
  Milk1.Set_Position 0, 100, 50
  
  'if texture creation is omitted by NemoX load it manualy
  'by
  'milk1.Set_Texture "your texturefile",your groupID if mesh has manygroups

End Sub




'we will used that sub for the engine initialization

Sub InitEngine()

 'first thing allocate memory for the main Object
  
  Set Nemo = New NemoX
  
  'check for the good version
 If NemoX.Get_EngineVersion < 0.072 Then
   MsgBox "NemoX version 1.072 is required as minimum" + Chr(10) + " Download it at http://perso.wanadoo.fr/malakoff/NemoXsetup.exe " + Chr(10) + " Or go to www.nemox.fr.st", vbInformation
   End
 End If
  
  
  
  Set Tool = New cNemo_Tools
  
  
  
   'allocate memory for our meshbuilder
    Set Mesh = New cNemo_Mesh
    
    '====New code======
    '
    '.......MEMORY ALLOCATION....
    Set KAMERA = New cNemo_Camera
    Set KEY = New cNemo_Input
    Set Milk1 = New cNemo_MSD3mesh
    Set Milk2 = New cNemo_MSD3mesh 'you can add your own model here
    Set Milk3 = New cNemo_MSD3mesh 'you can add your own model here
    
    
    

'we use this method
'now we allow the user to choose options
'32/16 bit backbuffer
  
  If Not (Nemo.INIT_ShowDeviceDLG(Form1.hWnd)) Then
   End 'terminate here if error
  End If
  
  'Nemo.Initialize Me.hWnd
  
  'set the back clearcolor
  Nemo.BackBuffer_ClearCOLOR = RGB(125, 125, 125) 'Gray
  
  
  
  'set some parameters
  Nemo.Set_ViewFrustum 10, 5500, 3.14 / 4, 1.01
  Nemo.Set_light True
  
  'set our camera
  KAMERA.Set_YRotation 45
  KAMERA.Set_EYE Tool.Vector(480, 50, -450)
  
  
 

 
End Sub


'=========================================
'  IN THIS sub we will handle with collision
'  detection
'
'
'=======================================
Sub CheckCollision()
    
    Dim Dest As D3DVECTOR
    
    'if any collision with our world we slide along
    If Mesh.CheckCollisionSliding(KAMERA.Get_Position, Dest, 40) Then
       KAMERA.Set_EYE Dest
    End If

 
     'if any collision with our Milkshape3D mesh we slide along

     If Milk1.Get_ColisionSliding(KAMERA.Get_Position, Dest, 20) Then
              KAMERA.Set_EYE Dest
    
     End If
 
 
     
 
  
 

End Sub












'this sub is the main loop for a game or 3d apllication
Sub gameLoop()

            'loop untill player press 'ESCAPE'
      'Nemo.Set_CullMode D3DCULL_NONE
      Nemo.Set_EngineRenderState D3DRS_ZENABLE, 1
    Do
       ZZ = ZZ + 3.14 / 50
       If ZZ > 3.14 * 2 Then ZZ = 0
               '=====Keyboard handler can be added here
               Call GetKey
               DoEvents
               
               CheckCollision
               
               
             
               'start the 3d renderer
               Nemo.Begin3D
                     '===============ADD game rendering mrthod here
               
               'draw our ground here
              
               Mesh.Render
               
               
              
               
               'render 3ds mesh
               Milk1.Render True
               
               
             
             
             
             
              
               
               'show the FPS at pixel(5,10) color White
               Nemo.Draw_Text "FPS:" + Str(Nemo.Framesperseconde), 5, 10, &HFFFF0000
               'Nemo.Draw_Text "Total Triangle Drawn=" + Str(Nemo.Get_EngineTotalTrianglesDrawn), 5, 25, &HFFFFFFFF
              
               
               Nemo.End3D
               'end the 3d renderer
            
            'check the player keyPressed
    Loop Until KEY.Get_KeyBoardKeyPressed(NEMO_KEY_ESCAPE)
            
            
    Call EndGame

End Sub



'----------------------------------------
'Name: GetKey
'----------------------------------------
Sub GetKey()



   'just Rotate left and Right
If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_LEFT) Then _
   KAMERA.Turn_Left (1.5 / 50) * RotationSpeed
If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_RIGHT) Then _
   KAMERA.Turn_Right (1.5 / 50) * RotationSpeed
    
    
    'just move Forward and Backward
    If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_UP) Then _
   KAMERA.Move_Foward 1 * MoveSpeed
    If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_RCONTROL) Then _
   KAMERA.Move_Foward 4 * MoveSpeed
   If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_DOWN) Then _
   KAMERA.Move_Backward 1 * MoveSpeed


   'just move uP and Down
  If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_ADD) Then _
   KAMERA.Strafe_UP 1
  If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_SUBTRACT) Then _
   KAMERA.Strafe_DOWN 1


      'to rotate UP/DOWN
     If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_NUMPAD8) Then
       KAMERA.Turn_UP (1 / 50) * RotationSpeed
    End If
    If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_NUMPAD2) Then _
       KAMERA.Turn_DOWN (1 / 50) * RotationSpeed
        
        'to move 0,8,-8
    If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_SPACE) Then _
    KAMERA.Set_Position Vector(0#, 8#, -408#), _
                                    Vector(0#, 8#, 500#)
    
    'to take a snapshot
    If KEY.Get_KeyBoardKeyPressed(NEMO_KEY_S) Then _
     Nemo.Take_SnapShot App.Path + "\Shot.bmp"




End Sub



Sub EndGame()
 'end of the demo
           Set KEY = Nothing
           Set KAMERA = Nothing
           Set Milk = Nothing
           Set Mesh = Nothing
           
            Nemo.Free  'free resources used by the engine
           Set Nemo = Nothing
            End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   EndGame
End Sub
