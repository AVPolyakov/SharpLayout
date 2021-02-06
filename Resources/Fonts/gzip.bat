for %%f in (*.ttf) do (
    "C:\Program Files\7-Zip\7z.exe" a -tgzip "%%f.gz" "%%f"
)