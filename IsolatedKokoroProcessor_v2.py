import wx
import sys
import spacy
from os import path, listdir, makedirs, rename, environ, pathsep
import multiprocessing
import importlib.util
from multiprocessing import Process, set_start_method, freeze_support

# import imageio_ffmpeg
# ffmpeg_path = imageio_ffmpeg.get_ffmpeg_exe()


def setup_ffmpeg():
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = path.dirname(path.abspath(__file__))

    # The folder where you put ffmpeg.exe, ffprobe.exe, ffplay.exe
    ffmpeg_dir = path.join(base_path, "ffmpeg", "bin")
    current_path = environ.get("PATH", "")
    new_path = f"{current_path}{pathsep}{ffmpeg_dir}"
    environ["PATH"] = new_path


def get_model_path(mname):
    # If running in PyInstaller bundle
    if getattr(sys, "frozen", False):
        base_path = sys._MEIPASS
    else:
        base_path = path.abspath(".")
    # environ['KOKORO_MODEL_DIR'] = path.join(base_path, 'models')
    # Point to the folder containing your .pth files
    return path.join(base_path, mname)


setup_ffmpeg()

# def get_kmodel_path():
#     if hasattr(sys, '_MEIPASS'):
#         # Navigate to the folder containing config.cfg you found earlier
#         return os.path.join(sys._MEIPASS, "en_core_web_sm", "en_core_web_sm-3.8.0")
#     return "en_core_web_sm"
# model_pth = get_kmodel_path()

# def force_load_spacy_model():
#     if hasattr(sys, '_MEIPASS'):
#         # 1. Locate the actual model folder (ensure this matches your _internal structure)
#         model_dir = os.path.join(sys._MEIPASS, "en_core_web_sm", "en_core_web_sm-3.8.0")

#         # 2. Manually load it as a module named 'en_core_web_sm'
#         spec = importlib.util.spec_from_file_location("en_core_web_sm", os.path.join(model_dir, "__init__.py"))
#         model_module = importlib.util.module_from_spec(spec)
#         sys.modules["en_core_web_sm"] = model_module
#         spec.loader.exec_module(model_module)

#         return "en_core_web_sm" # Return the string name, which now works!
#     return "en_core_web_sm"


def load_spacy_model():
    if hasattr(sys, "_MEIPASS"):
        # Start at the root of your model collection
        base_search_path = path.join(sys._MEIPASS, "en_core_web_sm")

        # 1. Check if the config is right here
        if path.exists(path.join(base_search_path, "config.cfg")):
            return base_search_path

        # 2. Check one level deeper (for that 3.8.0 folder)
        if path.exists(base_search_path):
            for folder in listdir(base_search_path):
                subfolder = path.join(base_search_path, folder)
                if path.isdir(subfolder) and "config.cfg" in listdir(subfolder):
                    return subfolder

    return "en_core_web_sm"  # Dev fallback


# Use it before KPipeline
# model_pth = force_load_spacy_model()

# spaxypathtop = get_model_path("en_core_web_sm")
# spaxypath = os.path.join(spaxypathtop, "en_core_web_sm-3.8.0")
# print(spaxypath)

# # Force the pipeline to use your bundled file
# model_pth = get_res("kokoro-v1_0.pth")


# pipeline = KPipeline(lang_code='a', model=model_path)

# environ["PATH"] += pathsep + r'C:\Users\ramdi\KokoroEnv\ffmpeg-2026-03-01-git-862338fe31-full_build\ffmpeg-2026-03-01-git-862338fe31-full_build\bin'
# environ["IMAGEIO_FFMPEG_EXE"] = ffmpeg_exe
# environ["IMAGEIO_FFMPEG_EXE"] = r"C:\Users\ramdi\KokoroEnv\ffmpeg-2026-03-01-git-862338fe31-full_build\ffmpeg-2026-03-01-git-862338fe31-full_build\bin\ffmpeg.exe"


import torch

torch.serialization.add_safe_globals(
    ["omegaconf.listconfig.ListConfig", "omegaconf.dictconfig.DictConfig"]
)

import pickle
import shutil
from natsort import natsorted
import whisperx
import subprocess
import numpy as np
import faster_whisper
import ctranslate2
from re import findall, sub
import soundfile as sf
from pathlib import Path
from transformers import Wav2Vec2ForCTC, Wav2Vec2Processor
from kokoro.pipeline import KPipeline
from threading import Thread
from pandas import read_excel, read_pickle, DataFrame, read_csv
from moviepy import concatenate_videoclips, concatenate_audioclips
from moviepy import VideoFileClip, AudioFileClip, CompositeAudioClip

from wxSlides import wxPresentation, wxTextBox


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
def SafeLog(message):
    if wx.IsMainThread():
        wx.LogMessage(message)
    else:
        wx.CallAfter(wx.LogMessage, message)


def ComponentWriter(fpath, name, comps):
    with open(fpath + name, "wb") as f:
        pickle.dump(comps, f)


def ComponentReader(fpath, name):
    with open(fpath + name, "rb") as f:
        comps = pickle.load(f)
    return comps


def StandardizedExcelReader(fpath):
    Files = [i for i in listdir(fpath) if ".xlsm" in i or ".xlsx" in i]
    # print(Files)
    FullSheet = read_excel(fpath + Files[0], sheet_name=None)
    SheetData = FullSheet.items()
    keys, dfs = [[] for i in range(2)]
    for key, df in SheetData:
        keys.append(key)
        dfs.append(df)

    StepKeys, StepPIDs, StepTools = [[] for i in range(3)]
    count = 0
    for i in dfs:
        if len(i["Item number"]) > 0:
            mask = i["Item number"] == "Tool"
            i["grouper"] = mask.cumsum()
            ig = {group_key: group_df for group_key, group_df in i.groupby("grouper")}
            ig[1].columns = ig[1].iloc[0]
            ig[1] = ig[1].loc[:, :"Quantity"]
            StepKeys.append(keys[count])
            StepPIDs.append(
                ig[0].dropna().reset_index(drop=True).drop(["grouper"], axis=1)
            )
            StepTools.append(
                ig[1]["Tool Description"].dropna().reset_index(drop=True)[1:]
            )

        count += 1
    return keys, StepPIDs, StepTools


# ~~~~~~~~~~~~~~~~~~~~~~~~~~~~


class Thor(wx.Frame):
    def __init__(self, parent, id, nlp):
        super().__init__(
            parent,
            id,
            "Thorlabs Instruction Transcription Interface",
            size=(1500, 1500),
        )
        self.nlp = nlp

        # wx.Frame.__init__(self,parent,id,'Thorlabs Origin Interface',size=(750,500))

        self.panel = wx.Panel(self)
        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        self.presMaker = wxPresentation(self.panel)

        # Displaying Text
        # font = wx.Font(10, wx.DECORATIVE, wx.NORMAL, wx.BOLD)
        buttonSizer = wx.StaticBoxSizer(wx.VERTICAL, self.panel)

        self.BOMWriterCB = wx.ComboBox(
            buttonSizer.StaticBox,
            wx.ID_ANY,
            choices=["No BOM", "BOM"],
            style=wx.CB_READONLY,
        )
        self.BOMWriterCB.SetValue("No BOM")
        buttonSizer.Add(
            self.BOMWriterCB, wx.SizerFlags(0).Align(wx.TOP).Border(wx.ALL, 10)
        )

        CPS = wx.Button(buttonSizer.StaticBox, label="Core Path")
        buttonSizer.Add(CPS, wx.SizerFlags(0).Align(wx.TOP).Border(wx.ALL, 10))

        self.VideoCombButton = wx.Button(
            buttonSizer.StaticBox, label="Video Compressor"
        )
        buttonSizer.Add(
            self.VideoCombButton, wx.SizerFlags(0).Align(wx.TOP).Border(wx.ALL, 10)
        )

        self.VideoSliceButton = wx.Button(buttonSizer.StaticBox, label="Video Slicer")
        buttonSizer.Add(
            self.VideoSliceButton, wx.SizerFlags(0).Align(wx.TOP).Border(wx.ALL, 10)
        )

        # self.LCB = wx.StaticText(buttonSizer.StaticBox, -1, "Audio Transcription")
        # self.AudioWriterCB = wx.ComboBox(
        #    buttonSizer.StaticBox,
        #    wx.ID_ANY,
        #    choices=["Original", "Rewrite"],
        #    style=wx.CB_READONLY,
        # )
        # self.AudioWriterCB.SetValue("Original")
        # buttonSizer.Add(self.LCB, wx.SizerFlags(0).Align(wx.TOP).Border(wx.ALL, 10))
        # buttonSizer.Add(
        #    self.AudioWriterCB, wx.SizerFlags(0).Align(wx.TOP).Border(wx.ALL, 10)
        # )
        # self.LCB.Hide()
        # self.AudioWriterCB.Hide()

        mainSizer = wx.BoxSizer(wx.VERTICAL)
        middleSizer = wx.BoxSizer(wx.HORIZONTAL)
        middleSizer.Add(buttonSizer, 0, wx.EXPAND | wx.TOP | wx.RIGHT, 10)
        middleSizer.Add(self.presMaker, 1, wx.EXPAND | wx.TOP | wx.RIGHT, 10)
        mainSizer.Add(middleSizer, 1, wx.EXPAND | wx.ALL, 10)

        self.saveFileTBox = wxTextBox(self.panel, title="File Name")
        self.saveFileTBox.Text = "example_presentation"
        self.saveFileButton = wx.Button(self.saveFileTBox.StaticBox, label="Save")
        self.saveFileButton.Disable()
        self.saveFileTBox.Add(self.saveFileButton, wx.SizerFlags(0).Border())
        mainSizer.Add(self.saveFileTBox, wx.SizerFlags(0).Expand().Border())

        self.logCtrl = wx.TextCtrl(self.panel, style=wx.TE_MULTILINE | wx.TE_READONLY)
        self.logCtrl.Hide()
        mainSizer.Add(self.logCtrl, 1, wx.EXPAND | wx.ALL, 10)

        self.showLogButton = wx.ToggleButton(self.panel, label="Show Log")
        mainSizer.Add(
            self.showLogButton,
            wx.SizerFlags(0).Align(wx.BOTTOM | wx.LEFT).Border(wx.ALL, 10),
        )

        self.panel.SetSizer(mainSizer)

        self.logger = wx.LogTextCtrl(self.logCtrl)
        wx.Log.SetActiveTarget(self.logger)

        # self.Bind(wx.EVT_COMBOBOX, self.on_combo_selection, self.AudioWriterCB)
        # self.Bind(wx.EVT_COMBOBOX, self.BOMSelection, self.BOMWriterCB)
        self.Bind(wx.EVT_BUTTON, self.PathSelector, CPS)
        self.Bind(wx.EVT_BUTTON, self.VideoCombination, self.VideoCombButton)
        self.Bind(wx.EVT_BUTTON, self.TranscriptionModel, self.VideoSliceButton)
        self.Bind(wx.EVT_TOGGLEBUTTON, self.OnToggleLog, self.showLogButton)
        self.Bind(wx.EVT_BUTTON, self.OnSavePPTX, self.saveFileButton)
        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~

        self.CorePath = None
        self.StructurePath = None
        self.AudioPath = None
        self.VideoPath = None
        self.TextPath = None
        self.SegPath = None

        self.WordSegments = None
        self.StepAudio = None
        self.StepVideo = None
        self.VideoClips = None
        self.ChangeLines = None
        self.VideoSlicerButton = None
        self.BKeys = None
        self.StepPIDs = None
        self.StepTools = None

        status = self.CreateStatusBar()
        status_font = self.GetStatusBar().GetFont()

        new_font_size = status_font.GetPointSize() + 2  # Increase size by 2 points
        new_font = wx.Font(
            new_font_size,
            status_font.GetFamily(),
            status_font.GetStyle(),
            status_font.GetWeight(),
            status_font.GetUnderlined(),
            status_font.GetFaceName(),
        )
        self.GetStatusBar().SetFont(new_font)
        self.Show()

    # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    def OnToggleLog(self, event):
        doShow = self.showLogButton.GetValue()
        if doShow:
            self.logCtrl.Show()
            self.showLogButton.SetLabel("Hide Log")
        else:
            self.logCtrl.Hide()
            self.showLogButton.SetLabel("Show Log")
        self.panel.Layout()

    def OnSavePPTX(self, event):
        savePath = self.CorePath + self.saveFileTBox.Text + ".pptx"
        wx.LogMessage(f"Saving generated slides to {savePath}")
        self.presMaker.Save(savePath)

    # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~

    def closebutton(self, event):
        self.Close(True)

    def closewindow(self, event):
        self.Destroy()

    def PathSelector(self, e):

        fileDialog = wx.DirDialog(
            frame, "Choose a directory", "", wx.DD_DEFAULT_STYLE | wx.DD_DIR_MUST_EXIST
        )
        if fileDialog.ShowModal() == wx.ID_OK:
            result = fileDialog.GetPath().replace("\\", "/") + "/"
            self.CorePath = result
            Structure = listdir(result)
            self.status = self.SetStatusText(
                "Successfully loaded: " + result + " -- Structure: " + str(Structure)
            )
        else:
            self.status = self.SetStatusText(
                "Command Terminated."
            )  # Untoggle the button
        #                 print(self.filepath)
        if "BOM" not in str(Structure):
            makedirs((result + "BOM/"), exist_ok=True)
            self.StructurePath = result + "BOM/"
            self.FileMover()
            self.StructurePath = None
        else:
            self.StructurePath = result + "BOM/"
            self.FileMover()
            self.StructurePath = None
        if "Videos" not in str(Structure):
            makedirs((result + "Videos/"), exist_ok=True)
            self.StructurePath = result + "Videos/"
            self.FileMover()
            self.StructurePath = None
        else:
            self.StructurePath = result + "Videos/"
            self.FileMover()
            self.StructurePath = None
        if "StepSegs" not in str(Structure):
            makedirs((result + "StepSegs/"), exist_ok=True)
        if "StepSegsAudio" not in str(Structure):
            makedirs((result + "StepSegsAudio/"), exist_ok=True)
        if "StepSegsTxt" not in str(Structure):
            makedirs((result + "StepSegsTxt/"), exist_ok=True)

        self.TextPath = self.CorePath + "StepSegsTxt/"
        self.VideoPath = self.CorePath + "Videos/"
        self.SegPath = self.CorePath + "StepSegs/"
        self.AudioPath = self.CorePath + "StepSegsAudio/"
        # self.AudioWriterCB.Show()
        # self.LCB.Show()
        self.Layout()

    def FileMover(self):
        source_path = self.CorePath
        destination_path = self.StructurePath
        if "BOM" in self.StructurePath:
            CoreFiles = [i for i in listdir(source_path) if ".xlsm" in i]
        if "Videos" in self.StructurePath:
            CoreFiles = [
                i for i in listdir(source_path) if ".MP4".casefold() in i.casefold()
            ]
        if len(CoreFiles) > 0:
            # try:
            # Move the files
            for i in CoreFiles:
                print(str(source_path) + "/" + str(i), str(destination_path))
                shutil.move(str(source_path) + "/" + str(i), str(destination_path))
                self.status = self.SetStatusText(
                    "Moved files into -- "
                    + str(destination_path)
                    + str(CoreFiles[0])
                    + " successfully."
                )  # Untoggle the button

        else:
            self.status = self.SetStatusText(
                "Nothing to move, or Error moving file into -- "
                + str(destination_path)
                + ". Verify it exists in core directory or if moved earlier."
            )  # Untoggle the button

    def VideoCombination(self, e):
        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        wx.LogMessage("Begin video combination/renaming.")

        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        # logger = (self.statusbar)
        # Corepath = self.CorePath
        # VideoPath=Corepath+'Videos/';self.VideoPath = VideoPath;self.SegPath = self.CorePath+'StepSegs/'#; self.AudioPath = Corepath+'StepSegsAudio/'
        VideoPath = self.VideoPath
        self.VideoCombButton.Disable()
        self.Layout()

        def work():
            try:
                if "combined.mp4" not in listdir(VideoPath):
                    VFiles = [
                        i
                        for i in listdir(VideoPath)
                        if ".MP4".casefold() in i.casefold()
                        if ".jpg" not in i
                    ]
                    # print(VFiles)
                    if len(VFiles) > 1:
                        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        wx.CallAfter(
                            wx.LogMessage,
                            "Multiple videos identified. Initiate sorting sequence.",
                        )
                        # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                        SegVidInds = natsorted(VFiles)
                        #                 SegVidInds = np.argsort(np.concatenate([np.double(findall(r'\d+', i.replace('.MP4'.casefold(),'').replace(Corepath+"Videos/",''))) for i in VFiles]))
                        #                 SortedVFiles = [VideoFileClip(VideoPath+VFiles[SegVidInds[i]]) for i in range(len(SegVidInds))]
                        SortedVFiles = [
                            VideoFileClip(VideoPath + i) for i in SegVidInds
                        ]
                        SortedVFiles[0].save_frame(
                            self.CorePath + "StepSegs/" + "FirstFrame.jpg", t=0
                        )
                        wx.CallAfter(
                            wx.LogMessage, "Sorting complete. Initiate concatenation."
                        )

                        CombinedVid = concatenate_videoclips(SortedVFiles)
                        wx.CallAfter(
                            wx.LogMessage,
                            "Concatenation complete. Writing video to output. Please be patient.",
                        )

                        CombinedVid.write_videofile(
                            VideoPath + "combined.mp4",
                            temp_audiofile="temp-audio.m4a",
                            remove_temp=True,
                            audio_codec="aac",
                            codec="libx264",
                            threads=8,
                            logger=None,
                            preset="veryfast",
                        )
                        # CombinedVid.write_videofile(Corepath+'/Videos/combined.mp4',temp_audiofile="temp-audio.m4a", remove_temp=True, audio_codec="aac", codec="libx264",logger=logger)
                        # wx.CallAfter(self.statusbar.SetStatusText, "Export Complete!")
                        # video = VideoFileClip(VideoPath+'combined.mp4')
                        wx.CallAfter(
                            wx.LogMessage,
                            "Video writing complete. Now writing audio to output.",
                        )

                        CombinedVid.audio.write_audiofile(
                            self.CorePath + "StepSegsAudio/" + "combined.mp3"
                        )

                        #                         CombinedVid.audio.write_audiofile(self.CorePath+'StepSegsAudio/'+"wavf_subs.wav", fps=16000, nbytes=2, codec='pcm_s16le', logger=None)
                        # self.AudioPath = self.CorePath+'StepSegsAudio/'
                        CombinedVid.audio.close()
                        CombinedVid.close()
                    if len(VFiles) == 1:
                        wx.CallAfter(
                            wx.LogMessage,
                            "Single video identified. Initiate renaming logic.",
                        )
                        rename(VideoPath + VFiles[0], VideoPath + "combined.mp4")
                        video = VideoFileClip(VideoPath + "combined.mp4")
                        video.save_frame(
                            self.CorePath + "StepSegs/" + "FirstFrame.jpg", t=0
                        )
                        wx.CallAfter(wx.LogMessage, "Audio embedding start.")
                        video.audio.write_audiofile(
                            self.CorePath + "StepSegsAudio/" + "combined.mp3"
                        )
                        #                 video.audio.write_audiofile(self.CorePath+'StepSegsAudio/'+"combined.wav", fps=16000, nbytes=2, codec='pcm_s16le', logger=None)
                        #                         CombinedVid.audio.write_audiofile(self.CorePath+'StepSegsAudio/'+"wavf_subs.wav", fps=16000, nbytes=2, codec='pcm_s16le', logger=None)
                        video.audio.close()
                        video.close()
                    if len(VFiles) == 0:
                        wx.CallAfter(wx.LogMessage, "No videos avaialable in path :(")

                        # self.status = self.SetStatusText('No videos avaiable in path :(')  # Untoggle the button
                else:
                    video = VideoFileClip(VideoPath + "combined.mp4")
                    video.save_frame(
                        self.CorePath + "StepSegs/" + "FirstFrame.jpg", t=0
                    )
                    video.close()
                    wx.CallAfter(
                        wx.LogMessage,
                        "combined.mp4 already exists. Please delete or move if another instance needs to be created.",
                    )

                    # self.status = self.SetStatusText('combined.mp4 already exists. Please delete or move if another instance needs to be created.')  # Untoggle the button
                    wx.CallAfter(
                        wx.LogMessage,
                        "Rewriting audio of existing combined.mp4 file.",
                    )
                    video = VideoFileClip(VideoPath + "combined.mp4")
                    video.audio.write_audiofile(
                        self.CorePath + "StepSegsAudio/" + "combined.mp3"
                    )
                wx.CallAfter(
                    wx.LogMessage,
                    "Audio successfully embedded. Please proceed to next step.",
                )

            except Exception as e:
                # wx.CallAfter(wx.LogMessage, f"Error: {(AudioFile+'combined.mp3')}")
                wx.CallAfter(wx.LogMessage, f"Error: {str(e)}")
                wx.CallAfter(self.VideoSliceButton.Enable)

        Thread(target=work, daemon=True).start()
        self.VideoCombButton.Enable()
        self.Layout()

    def TimeSlicer(self):
        word_segments = self.WordSegments
        StartSeg = -1
        EndSeg = -1
        Step = []
        for i in range(len(word_segments)):
            if i + 2 < len(word_segments):
                if (
                    "start" in word_segments[i]["word"].casefold()
                    and "step" in word_segments[i + 1]["word"].casefold()
                ):
                    StartSeg = i
                if (
                    "end" in word_segments[i]["word"].casefold()
                    or "and" in word_segments[i]["word"].casefold()
                    or "finish" in word_segments[i]["word"].casefold()
                    or "stop" in word_segments[i]["word"].casefold()
                ) and "step" in word_segments[i + 1]["word"].casefold():
                    EndSeg = i + 3
            if StartSeg != -1 and EndSeg != -1:
                Step.append(word_segments[StartSeg:EndSeg])
                StartSeg = -1
                EndSeg = -1
        StichedStep = [" ".join([j["word"] for j in i]) for i in Step]
        # print('End Stitching. Attempting any corrections')
        for i in StichedStep:
            if "and step" in i.casefold() or "And step" in i.casefold():
                i.casefold().replace("and step", "end step")
        # print('Corrections successful!')
        CompleteStepTiming = [
            [Step[i][0]["start"], Step[i][len(Step[i]) - 1]["end"]]
            for i in range(len(Step))
        ]
        return Step, StichedStep, CompleteStepTiming

    def TranscriptionModel(self, e):
        self.status = self.SetStatusText(
            "Transcription sequence initiated."
        )  # Untoggle the button
        #         AudioFile = self.AudioPath+'combined.mp3'
        # self.AudioPath = self.CorePath+'StepSegsAudio/';
        AudioFile = self.AudioPath
        #         if torch.cuda.is_available():
        #             device = "cuda"
        #         else:
        #             device='cpu'
        device = "cpu"
        batch_size = 8
        if AudioFile == "":
            raise Exception("Please enter valid filepath")
        # if batch_size <16:
        #     compute_type = "int8"
        # else:
        #     compute_type = "float16"
        compute_type = "int8"
        self.VideoSliceButton.Disable()
        self.Layout()

        self.status = self.SetStatusText("Initiating thread...")  # Untoggle the button

        wx.LogMessage("Loading & Transcribing... " + str(device))

        def work():
            try:
                AudioFile = self.AudioPath
                device = "cpu"
                compute_type = "int8"
                batch_size = 1
                wx.CallAfter(wx.LogMessage, "Attempting to load large-v3 model.")
                Model = whisperx.load_model(
                    "large-v3", device, compute_type=compute_type
                )
                wx.CallAfter(wx.LogMessage, "Attempting to load audio.")
                LoadedAudio = whisperx.load_audio(Path(AudioFile + "combined.mp3"))
                wx.CallAfter(wx.LogMessage, "Initiate transcription.")
                PreliminaryTranscription = Model.transcribe(
                    LoadedAudio, batch_size=batch_size, language="en"
                )
                wx.CallAfter(wx.LogMessage, "Time projection.")
                AlignModel, Metadata = whisperx.load_align_model(
                    language_code="en", device=device
                )
                AlignedTranscription = whisperx.align(
                    PreliminaryTranscription["segments"],
                    AlignModel,
                    Metadata,
                    AudioFile + "combined.mp3",
                    device,
                    return_char_alignments=False,
                )
                wx.CallAfter(wx.LogMessage, "Begin writing.")
                self.WordSegments = AlignedTranscription["word_segments"]
                wx.CallAfter(
                    wx.LogMessage,
                    "Alignment sequence complete. Initiate step isolation sequence.",
                )
                self.FullSteps, self.StichedSteps, self.StichedTiming = (
                    self.TimeSlicer()
                )
                wx.CallAfter(
                    wx.LogMessage,
                    "Step Isolation sequence complete. Obtaining audio slices",
                )
                # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                self.StepAudio = [
                    self.AudioWriter(
                        self.StichedSteps[i],
                        self.AudioPath + "AFHeart",
                        i,
                    )
                    for i in range(len(self.StichedSteps))
                ]
                # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                wx.CallAfter(
                    wx.LogMessage, "Audio slices obtained. Initiate video rendering."
                )
                self.StepVideo = [
                    self.VideoStepWriter(self.StichedTiming[i], i)
                    for i in range(len(self.StichedTiming))
                ]
                wx.CallAfter(
                    wx.LogMessage,
                    "Videos successfully rendered. Presentation creation initiated.",
                )
                if self.BOMWriterCB.GetValue() == "BOM":
                    wx.CallAfter(wx.LogMessage, "Extracting BOM data.")
                    self.BOMWriter()
                    wx.CallAfter(
                        wx.LogMessage, "BOM data extracted. Starting presentation."
                    )
                # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
                wx.CallAfter(wx.LogMessage, "Starting presentation.")
                wx.CallAfter(self.AddSlides)
                # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~

                wx.CallAfter(
                    wx.LogMessage,
                    "KokoroAI Processes Complete. Thank you for your patience!",
                )  # Thread-safe update

            except Exception as e:
                # wx.CallAfter(wx.LogMessage, f"Error: {(AudioFile+'combined.mp3')}")
                wx.CallAfter(wx.LogMessage, f"Error: {str(e)}")
                wx.CallAfter(self.VideoSliceButton.Enable)

        # 2. Fire it off in a thread immediately
        Thread(target=work, daemon=True).start()
        # wx.LogMessage("Processing...")
        #         self.StepVideo = [self.VideoStepWriter(self.StichedTiming[i],i) for i in range(len(self.StichedTiming))]
        #         self.status = self.SetStatusText('Excel Gener.')  # Untoggle the button
        self.VideoSliceButton.Enable()
        self.Layout()

    def AudioWriter(self, StepSegments, Tag, Index):
        pipeline = self.nlp
        generator = pipeline(
            StepSegments,
            voice="af_heart",  # <= change voice here
            speed=0.8,
            split_pattern=r"\n+",
        )

        AudioClips = []
        TextClips = []
        for i, (gs, ps, audio) in enumerate(generator):
            AudioClips.append(Tag + "Step" + str(i) + ".mp3")
            sf.write(
                Tag + "Step" + str(i) + ".mp3", audio, 24000
            )  # save each audio file
        FullAudio = concatenate_audioclips([AudioFileClip(i) for i in AudioClips])
        FullAudio.write_audiofile(Tag + "FullStep" + str(Index) + ".mp3", 24000)
        return Tag + "FullStep" + str(Index) + ".mp3"

    def VideoStepWriter(self, e, Index):
        # print('Start Video Step Writing')
        VideoPath = self.VideoPath
        AudioPath = self.AudioPath
        Timing = self.StichedTiming
        VideoClips = []
        #         NoSubVideo = VideoFileClip(VideoPath+'combined.mp4').without_audio()
        Video = VideoFileClip(VideoPath + "combined.mp4").without_audio()
        #         Video = CompositeVideoClip([NoSubVideo, self.subs])

        MiniAud = AudioFileClip(self.StepAudio[Index])
        MiniVid = Video[Timing[Index][0] : Timing[Index][1]]
        MiniVid = MiniVid.with_audio(MiniAud)
        #         print('Video Sliced')
        #         print(self.SegPath+"AFHeart"+str(Index)+".mp4")
        MiniVid.write_videofile(
            self.SegPath + "AFHeart" + str(Index) + ".mp4",
            codec="libx264",
            audio_codec="aac",
            preset="veryfast",
            logger=None,
            threads=8,
        )

        VideoClips.append(self.SegPath + "AFHeart" + str(Index) + ".mp4")
        self.VideoClips = VideoClips

    # def AudioReWriter(self, Tag, Index):

    # def on_complete(self, segments):
    #    wx.LogMessage(f"Done! {len(segments)} segments found.")
    #    self.VideoSliceButton.Enable()

    #     def on_start_export(self, event):
    #         self.VideoSliceButton.Disable() # Prevent multiple clicks
    #         thread = threading.Thread(target=self.render_comb_video)
    #         thread.start()
    def render_comb_video(self, clip, output_path, threads):
        clip.write_videofile(
            output_path,
            temp_audiofile="temp-audio.m4a",
            remove_temp=True,
            audio_codec="aac",
            codec="libx264",
            threads=threads,
            preset="ultrafast",
            logger=None,
        )

    #         wx.CallAfter(self.on_export_finished)
    #     def on_export_finished(self):
    #         self.VideoSliceButton.Enable()
    #         wx.MessageBox("Export complete!", "Info")
    def BOMWriter(self):
        self.BOMPath = self.CorePath + "BOM/"
        BomPath = self.BOMPath
        self.BKeys, self.StepPIDs, self.StepTools = StandardizedExcelReader(BomPath)
        ComponentWriter(
            self.TextPath, "AFHeartTxt.pkl", [self.StepPIDs, self.StepTools]
        )

    # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    def AddSlides(self):
        SegInds = np.argsort(
            np.concatenate(
                [
                    np.double(
                        findall(
                            r"\d+", i.replace(".mp4", "").replace(self.SegPath, "")[-3:]
                        )
                    )
                    for i in listdir(self.SegPath)
                    if "AFHeart" in i
                ]
            )
        )

        FullStepVidPaths = [
            self.SegPath + i for i in listdir(self.SegPath) if ".jpg" not in i
        ]

        StepData = (
            read_pickle(self.TextPath + "AFHeartTxt.pkl")
            if self.BOMWriterCB.GetValue() == "BOM"
            else None
        )

        # StichedDF = DataFrame(self.StichedSteps).to_csv(
        #    self.TextPath + "Uncorrected_Steps.csv"
        # )
        # textFileName = "Uncorrected_Steps.csv"  # if self.AudioWriterCB.GetValue() == "Original" else "Corrected_Steps.csv"

        # textData = read_csv(self.TextPath + textFileName, encoding="latin1")[
        #    "0"
        # ].to_list()

        for i in range(len(FullStepVidPaths)):
            title = f"Step {i + 1}"
            print(SegInds[i])
            StepComponentData = StepData[0][i] if StepData is not None else None
            StepToolData = (
                DataFrame(list(StepData[1][i]), columns=["Tools"])
                if StepData is not None
                else None
            )

            vidFileName = (
                FullStepVidPaths[SegInds[i]]
                if SegInds[i] < len(FullStepVidPaths)
                else None
            )

            self.presMaker.AddStepSlide(
                title,
                self.StichedSteps[i],
                vidFileName,
                StepComponentData,
                StepToolData,
            )

        self.saveFileButton.Enable()

    # def on_combo_selection(self, event):
    #    # 4. Get the selected value
    #    selected_item = event.GetEventObject().GetValue()
    #    if selected_item == "Original":
    #        self.status = self.SetStatusText("Raw audio transcription being used")
    #    else:
    #        data = read_csv(self.TextPath + "Corrected_Files.csv", encoding="latin1")
    #        #             print(Changelines)
    #        self.ChangeLines = data["0"].to_list()
    #        #             self.ChangeLines = [sub(r" ?\[.*?\]","", item) for item in Changelines]

    #        #             print(self.ChangeLines)
    #        self.status = self.SetStatusText(
    #            "Rewrite requested. Please make sure to modify a file named Corrected_Files.csv, with the lines only occupying the same number of cells as the Uncorrected_Steps.csv file does."
    #        )

    # def BOMSelection(self, e):
    #    # 4. Get the selected value
    #    selected_item = e.GetEventObject().GetValue()
    #    return selected_item

    # ~~~~~~~~~~~~~~~~~~~~~~~~~~~~


if __name__ == "__main__":
    # REQUIRED: Must set 'spawn' before creating any processes (default on Windows)
    multiprocessing.freeze_support()

    try:
        set_start_method("spawn", force=True)
    except RuntimeError:
        pass
    try:
        # model_pth = force_load_spacy_model()
        model_pth = load_spacy_model()
        pipeline = KPipeline(lang_code="a", model=model_pth)
        print("Model and Pipeline loaded successfully!")
        # nlp_model = load_spacy_model()
        # print("Model loaded successfully!")
    except Exception as e:
        print(f"Failed to load model: {e}")
        nlp_model = None
    app = wx.App()
    frame = Thor(parent=None, id=-1, nlp=pipeline)
    frame.Show()
    app.MainLoop()
