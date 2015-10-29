# Audio snippets category
These snippets are intended to help managing Audio files capabilities.

# General Information about this repository
 - .snippet files contains the formatted code that can be loaded through Visual Studio's code editor context menu.
 - .vb files contains the raw code that can be copied then pasted in any project.
 - Each .snippet and .vb file contains a #Region section and/or Xml documentation with code examples.
 
Feel free to use and/or modify any file of this repository.

If you like the job I've done, then please contribute with improvements to these snippets or by adding new ones.

# Audio Util
Contains related audio utilities, such a WAV recorder, a WAV/MP3/MIN player, set the volume of the app or mute the system device.

Public Members Summary

 - Types
   - AudioUtil.AudioPlayer : IDisposable
   - AudioUtil.StereoVolume <Serializable>
   - AudioUtil.WaveRecorder : IDisposable

 - Cosntructors
   - AudioUtil.AudioPlayer.New()
   - AudioUtil.AudioPlayer.New(Form)
   - AudioUtil.StereoVolume(Integer, Integer)
   - AudioUtil.WaveRecorder.New()

 - Properties
   - AudioUtil.AudioPlayer.Filepath As String
   - AudioUtil.AudioPlayer.Status As PlayerState
   - AudioUtil.AudioPlayer.PlaybackMode As AudioPlayMode
   - AudioUtil.AudioPlayer.Channels As Integer
   - AudioUtil.AudioPlayer.Length As Integer
   - AudioUtil.AudioPlayer.Position As TimeSpan
   - AudioUtil.AudioPlayer.IsFileLoaded As Boolean
   - AudioUtil.StereoVolume.LeftChannel As Integer
   - AudioUtil.StereoVolume.RightChannel As Integer
   - AudioUtil.WaveRecorder.Status As AudioUtil.WaveRecorder.RecorderStatus

 - Enumerations
   - AudioUtil.ChannelMode As Integer
   - AudioUtil.AudioPlayer.PlayerState As Integer
   - AudioUtil.WaveRecorder.RecorderStatus As Integer

 - Functions
   - AudioUtil.GetAppVolume() As AudioUtil.StereoVolume

 - Methods
   - AudioUtil.MuteSystemVolume() 
   - AudioUtil.SetAppVolume(Integer) 
   - AudioUtil.SetAppVolume(Integer, Integer) 
   - AudioUtil.SetAppVolume(AudioUtil.StereoVolume) 
   - AudioUtil.AudioPlayer.LoadFile(String)
   - AudioUtil.AudioPlayer.UnloadFile
   - AudioUtil.AudioPlayer.Play(Opt: AudioPlayMode)
   - AudioUtil.AudioPlayer.Seek(Long)
   - AudioUtil.AudioPlayer.Seek(TimeSpan)
   - AudioUtil.AudioPlayer.Pause
   - AudioUtil.AudioPlayer.Resume
   - AudioUtil.AudioPlayer.Stop
   - AudioUtil.AudioPlayer.Dispose
   - AudioUtil.WaveRecorder.Record
   - AudioUtil.WaveRecorder.Stop
   - AudioUtil.WaveRecorder.Play
   - AudioUtil.WaveRecorder.Delete
   - AudioUtil.WaveRecorder.Save(String, Opt: Boolean)
   - AudioUtil.WaveRecorder.Dispose
