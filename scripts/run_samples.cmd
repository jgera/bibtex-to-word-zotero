@echo off
setlocal enabledelayedexpansion

echo ===================================================
echo     BibTeX to Zotero Word Reference Conversion
echo ===================================================

if "%~1"=="" goto run_all
if "%~1"=="1" goto run_sample1
if "%~1"=="2" goto run_sample2
if "%~1"=="3" goto run_sample3
if "%~1"=="all" goto run_all
goto run_custom

:run_all
echo [INFO] Running batch conversion for ALL samples...
python create_samples_and_organize.py
goto end

:run_sample1
echo [INFO] Converting Sample 1 (MiniLM)...
python bibtex_to_zotero_word.py --input samples\01_minilm\input.docx --bib samples\01_minilm\references.txt --output output\01_minilm_zotero.docx
goto end

:run_sample2
echo [INFO] Converting Sample 2 (NLP and LLM Survey)...
python bibtex_to_zotero_word.py --input samples\02_nlp_llm_survey\input.docx --bib samples\02_nlp_llm_survey\references.bib --output output\02_nlp_llm_zotero.docx
goto end

:run_sample3
echo [INFO] Converting Sample 3 (Multimodal Vision and RAG)...
python bibtex_to_zotero_word.py --input samples\03_cv_multimodal_rag\input.docx --bib samples\03_cv_multimodal_rag\references.bib --output output\03_cv_multimodal_zotero.docx
goto end

:run_custom
echo [INFO] Running custom parameters: %*
python bibtex_to_zotero_word.py %*
goto end

:end
echo ===================================================
echo [SUCCESS] Done! Generated files are saved in output\
echo ===================================================
