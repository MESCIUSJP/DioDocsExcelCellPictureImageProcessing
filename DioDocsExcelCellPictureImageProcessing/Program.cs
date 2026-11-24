using GrapeCity.Documents.Drawing;
using GrapeCity.Documents.Excel;
using GrapeCity.Documents.Imaging;
using GrapeCity.Documents.Text;
using System.Drawing;

// 入力ブックと出力ブックのパス
string inputPath = Path.Combine(AppContext.BaseDirectory, "test.xlsx");
string outputPath = Path.Combine(AppContext.BaseDirectory, "result.xlsx");

// 1. ワークブックを開く
var workbook = new Workbook();
workbook.Open(inputPath);
IWorksheet worksheet = workbook.Worksheets[0];

// 2. セル内画像（CellPicture）を含むすべてのセル範囲を取得
IRange pictureCells = worksheet.Cells.SpecialCells(SpecialCellType.CellPicture);

if (pictureCells != null)
{
    // 複数のセル範囲領域を順番に処理
    for (int areaIndex = 0; areaIndex < pictureCells.Areas.Count; areaIndex++)
    {
        IRange area = pictureCells.Areas[areaIndex];

        // 領域内の各行と各列から対象セルを取得
        for (int rowIndex = 0; rowIndex < area.RowCount; rowIndex++)
        {
            for (int columnIndex = 0; columnIndex < area.ColumnCount; columnIndex++)
            {
                // 現在のセルに設定されたセル内画像を取得
                IRange cell = area[rowIndex, columnIndex];
                CellPicture? cellPicture = cell.CellPicture;

                // 画像データがないセルは処理対象外
                if (cellPicture == null || cellPicture.ImageData == null)
                {
                    continue;
                }

                // 元の代替テキストを保持
                string altText = cellPicture.AltText;

                // 元の画像と編集後の画像をメモリストリームで保持
                using var originalImageStream = new MemoryStream(cellPicture.ImageData);
                using var processedImageStream = new MemoryStream();

                // 3. DioDocs for Imaging: 画像の編集（リサイズ・スタンプ描画・圧縮）
                ProcessImage(originalImageStream, processedImageStream);

                // 4. DioDocs for Excel: 編集済み画像と代替テキストで新しいCellPictureを設定
                byte[] processedBytes = processedImageStream.ToArray();
                cell.CellPicture = new CellPicture(processedBytes, altText);
            }
        }
    }
}

// 5. 編集後のワークブックを保存
workbook.Save(outputPath);

Console.WriteLine("セル内画像の最適化が完了しました。");

/// <summary>
/// 画像のリサイズ・テキスト描画・JPEG圧縮を行う
/// </summary>
static void ProcessImage(Stream inputStream, Stream outputStream)
{
    // 入力ストリームから元画像を読み込む
    using var srcBmp = new GcBitmap();
    srcBmp.Load(inputStream);

    // 最大幅800ピクセルを基準にリサイズ後のサイズを設定
    int targetWidth = srcBmp.PixelWidth;
    int targetHeight = srcBmp.PixelHeight;
    const int maxWidth = 800;

    // 横幅が上限を超える場合のみ、縦横比を維持して縮小
    if (targetWidth > maxWidth)
    {
        double ratio = (double)maxWidth / targetWidth;
        targetWidth = maxWidth;
        targetHeight = (int)(srcBmp.PixelHeight * ratio);
    }

    // 高品質バイキュービック補間でリサイズ
    using var resizedBmp = srcBmp.Resize(
        targetWidth,
        targetHeight,
        InterpolationMode.Cubic
    );

    // 画像に確認済テキスト（スタンプ）を描画
    using (GcBitmapGraphics g = resizedBmp.CreateGraphics())
    {
        var textFormat = new TextFormat
        {
            BackColor = Color.FromArgb(128, 255, 255, 255), // 半透明の白背景
            Font = Font.FromFile(GetSystemFontPath()), // 環境に応じたフォントを指定
            FontSize = 11,
            ForeColor = Color.Red
        };

        string stampText = $"確認済 {DateTime.Now:yyyy/MM/dd}";
        g.DrawString(
            stampText,
            textFormat,
            new PointF(10, 10)
        );
    }

    // JPEG形式（品質80%）で出力ストリームへ保存し容量を削減
    resizedBmp.SaveAsJpeg(outputStream, null, 80);
}

static string GetSystemFontPath()
{
    // Windows環境での日本語フォント例（Linux環境などで実行する場合は、フォントファイルのパスを環境に合わせて変更してください）
    return @"C:\Windows\Fonts\meiryo.ttc";
}
