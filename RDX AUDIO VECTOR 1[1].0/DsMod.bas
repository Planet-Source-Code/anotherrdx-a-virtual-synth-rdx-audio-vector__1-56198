Attribute VB_Name = "DsMod"
Option Explicit

Global DX               As New DirectX8
Global DS               As DirectSound8
Global DsB              As DirectSoundSecondaryBuffer8
Global dsbd             As DSBUFFERDESC
Global WaveFormat       As WAVEFORMATEX

Global FXe              As DirectSoundFXEcho8
Global FXr              As DirectSoundFXWavesReverb8

Global fxReverb         As DSFXWAVESREVERB
Global fxEcho           As DSFXECHO
Global DSEffects(2)     As DSEFFECTDESC
Global lresult(1)       As Long
'-----------------------------------------------------Variables application

Dim TmpBuffer()         As Integer      'Buffer ou est stockée le signal a copier

'-----------------------------------------------------Constantes
Const lSize             As Long = 2688  'Taille du buffer (toujours en OCTETS)
Const Amp               As Long = 32767 'Amplitude du signal
Sub SetDsFxEcho()
   
    With fxEcho
        .fFeedback = PotVal(14)
            If PotVal(13) = 0 Then
                .fWetDryMix = 0
                .fLeftDelay = 50
                .fRightDelay = 50
            Else
                .fLeftDelay = (PotVal(13) * 1000 * 4 / Tempo * 60 / 16) \ 1
                .fRightDelay = ((PotVal(13) + 2) * 1000 * 4 / Tempo * 60 / 16) \ 1
                .fWetDryMix = 50
            End If
        .lPanDelay = 0
    End With

    FXe.SetAllParameters fxEcho

End Sub

Sub SetDsFxReverb()
    
    With fxReverb
        .fInGain = 0
        .fReverbMix = -96 + (PotVal(12) * 96) \ 100
        .fReverbTime = PotVal(11) * 29 + 1
        .fHighFreqRTRatio = 0.999
    End With
    FXr.SetAllParameters fxReverb

End Sub

Sub beginFX()
    DsB.SetFX 2, DSEffects, lresult
    Set FXe = DsB.GetObjectinPath(DSFX_STANDARD_ECHO, 0, IID_DirectSoundFXEcho)
    Set FXr = DsB.GetObjectinPath(DSFX_STANDARD_WAVES_REVERB, 0, IID_DirectSoundFXWavesReverb)

    SetDsFxEcho
    SetDsFxReverb
End Sub

Sub UnloadDx()
    If Not (DsB Is Nothing) Then DsB.Stop
    Set DsB = Nothing
    Set DS = Nothing
    Set DX = Nothing
End Sub

Function InitDX(DSID As String) As Boolean
On Error GoTo FailedInit
    
    InitDX = True
    Set DS = DX.DirectSoundCreate(DSID)
    DS.SetCooperativeLevel Form1.hWnd, DSSCL_PRIORITY
    
    DSEffects(0).guidDSFXClass = DSFX_STANDARD_WAVES_REVERB
    DSEffects(1).guidDSFXClass = DSFX_STANDARD_ECHO
    
    With WaveFormat
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = 2
        .lSamplesPerSec = 44100
        .nBitsPerSample = 16
        .nBlockAlign = WaveFormat.nBitsPerSample / 8 * WaveFormat.nChannels
        .lAvgBytesPerSec = WaveFormat.lSamplesPerSec * WaveFormat.nBlockAlign
    End With
    
    With dsbd
        .fxFormat = WaveFormat
        .lFlags = DSBCAPS_CTRLFX Or DSBCAPS_CTRLVOLUME Or DSBCAPS_GETCURRENTPOSITION2
        .lBufferBytes = 1000000
    End With
    
    Set DsB = DS.CreateSoundBuffer(dsbd)
    
    'Cleaning Buffer BETA 1.0 by Dj-Wincha
        ReDim TmpBuffer(1000000 \ 2 - 1) As Integer
    Dim T As Variant
    'Rellena el Buffer (Español)
    'Fill the Buffer (English)
    For T = 0 To UBound(TmpBuffer) Step 2
        TmpBuffer(T) = ((T Mod (1000000 \ 2 - 1)) / (1000000 \ 2 - 1) * Amp) \ 1
        TmpBuffer(T + 1) = TmpBuffer(T)
    Next
    
    'Note:The param Start & Size is in OCTETS! (English) - Sorry for my bad english
    'Nota:Los Parametros Start & Size son en OCTETOS (Español)
    DsB.WriteBuffer 0, 1000000, TmpBuffer(0), DSBLOCK_DEFAULT
    Exit Function

FailedInit:
    InitDX = False
End Function
