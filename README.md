# geoguessr


    try{

        # プリンタの名前設定
        $printer = Get-WmiObject Win32_Printer | Where-Object Name -eq "EPSON PX-045A Series"
        $printer.SetDefaultPrinter() > $null

        # プリンタのカラー設定 $Trueならカラー、$Falseなら白黒（たぶん）
        Set-PrintConfiguration $printer.name -Color $False

        # ディレクトリ名入力
        $path = Read-Host "印刷対象のフォルダ名を入力してください"

        # ファイル一覧取得 日付フォルダがある場所を指定
        $absPath = Join-Path "C:\Users\user\Desktop\hena\" $path


        # 拡張子をカンマ区切りで指定
        $files = Get-ChildItem $absPath -Include *.doc,*.docx,*.pdf,*xlsx,*xls -Recurse
        # echo "印刷対象のファイルは以下です。" $files

        # ファイルごとに印刷

        $errorFlag = $false

        foreach ($file in $files){
            try{

                start-process -FilePath $file.fullName -Verb Print
                Write-Host "印刷に成功しました。ファイル名：" + $file.FullName -BackgroundColor Green

            }catch{

                $errorFlag = $true
                Write-Host "印刷できませんでした。ファイル名：" + $file.FullName -BackgroundColor Red

            }
        }

        if ($errorFlag.Equals($false)){
            Write-Host "すべての印刷が正常に終了しました。" -BackgroundColor Green
        } else {
            Write-Host "印刷できなかったファイルがあります。確認してください。" -BackgroundColor Red
        }

    } catch {

        Write-Error("エラーが発生しました。"+$_.Exception)

    }
