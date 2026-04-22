# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all
from PyInstaller.utils.hooks import copy_metadata
from PyInstaller.utils.hooks import collect_data_files,collect_submodules

datas = [('ffmpeg', 'ffmpeg'), ('kokoro-v1_0.pth', '.'), ('WhisperModel.bin', '.'), ('Kokoroconfig.json', '.'), ('Whisperconfig.json', '.'), ('voices', 'voices')]
binaries = []
hiddenimports = ['whisperx', 'whisper.asr', 'faster_whisper', 'faster_whisper.transcribe', 'omegaconf', 'torchaudio', 'whisperx.diarize', 'pyannote.audio', 'whisperx.alignment', 'pycparser.lextab', 'pycparser.yacctab', 'whisperx.transcribe', 'transformers.pipelines', 'transformers.pipelines.base', 'transformers.models', 'kokoro.model', 'ctranslate2.converters', 'kokoro.pipeline', 'misaki.g2p', 'transformers.models.wav2vec2.modeling_wav2vec2', 'transformers.models.wav2vec2.tokenization_wav2vec2', 'pyannote.audio.models.segmentation', 'pyannote.audio.pipelines.speaker_diarization', 'pyannote.audio.pipelines.utils', 'torch.utils.data.datapipes.datapatterns','spacy', 'spacy.kb', 'spacy.tokens', 'spacy.lang.en', 'en_core_web_sm', 'srsly','srsly.msgpack.util', 'catalogue', 'thinc.backends.win_ops','pydantic.deprecated.decorator','setuptools', 'pkg_resources', 'pkg_resources.extern', 'pkg_resources.py2_warn','packaging']

model_datas = collect_data_files('en_core_web_sm', include_py_files=True)
model_hidden = collect_submodules('en_core_web_sm')
spacy_hidden = collect_submodules('spacy')
setuptools_hidden = collect_submodules('setuptools')
pkg_hidden = collect_submodules('pkg_resources')
hiddenimports+= model_hidden
hiddenimports+= spacy_hidden
hiddenimports+= setuptools_hidden
hiddenimports+= pkg_hidden
datas += copy_metadata('imageio')
datas += copy_metadata('kokoro')	
datas += copy_metadata('ctranslate2')
datas += copy_metadata('transformers')
datas += copy_metadata('torch')
datas += copy_metadata('torchcodec')
datas += copy_metadata('misaki')
datas += copy_metadata('tqdm')
datas += copy_metadata('regex')
datas += copy_metadata('requests')
datas += copy_metadata('filelock')
datas += copy_metadata('numpy')
datas += copy_metadata('faster_whisper')
datas += copy_metadata('tokenizers')
datas += copy_metadata('whisperx')
datas += copy_metadata('imageio')
datas += copy_metadata('setuptools')
tmp_ret = collect_all('transformers')

datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('language_tags')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('espeakng_loader')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('ctranslate2')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('torchaudio')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('kokoro')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('misaki')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('onnxruntime')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('whisperx')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('faster_whisper')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('spacy')
datas += tmp_ret[0]; binaries += tmp_ret[1]; hiddenimports += tmp_ret[2]
tmp_ret = collect_all('en_core_web_sm')
datas += tmp_ret[0]

datas += collect_data_files('pyannote.audio')
datas += collect_data_files('whisperx')
datas += collect_data_files('spacy')
datas += model_datas

a = Analysis(
    ['IsolatedKokoroProcessor_v2.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TestKDir',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='TestKDir',
)
