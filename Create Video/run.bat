@echo off
echo ============================================================
echo  Create Video - PowerPoint to Narrated MP4
echo ============================================================
echo.

python create_video.py

echo.
if %ERRORLEVEL% NEQ 0 (
    echo Video creation failed with error code %ERRORLEVEL%.
) else (
    echo Video creation completed successfully.
)

pause
