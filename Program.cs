
using System.Reflection.Emit;
using NPOI.HSSF.Record;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;

const double RowHeight = 20;

Console.WriteLine("Tournify Gamesheet creator");


var games = new List<GameRecord>();
if (args.Length == 1)
{
    var filePath = args[0];

    using (var file = new FileStream(filePath, FileMode.Open, FileAccess.Read))
    {
        // Initialize the workbook
        IWorkbook workbook = new XSSFWorkbook(file);

        // Get the first sheet
        ISheet sheet = workbook.GetSheetAt(0);

        // Loop through each row in the sheet and read the data
        for (int i = 1; i <= sheet.LastRowNum; i++)
        {
            IRow row = sheet.GetRow(i);

            // Set the start and end times from the row data
            var startTimeString = row.GetCell(0)?.ToString() ?? string.Empty;
            var endTimeString = row.GetCell(1)?.ToString() ?? string.Empty;

            DateTime.TryParse(startTimeString, out DateTime startTime);
            DateTime.TryParse(endTimeString, out DateTime endTime);


            var referee1 = row.GetCell(9)?.ToString() ?? string.Empty;
            var referee2 = row.GetCell(10)?.ToString() ?? string.Empty;
            var scorer = row.GetCell(11)?.ToString() ?? string.Empty;

            if (referee2.StartsWith("Scorer:"))
            {
                scorer = referee2;
                referee2 = referee1;

            }

            var gamerecord = new GameRecord
            (
                StartTime: startTime,
                EndTime: endTime,
                Day: row.GetCell(2)?.ToString() ?? string.Empty,
                Field: row.GetCell(3)?.ToString() ?? string.Empty,
                Phase: row.GetCell(4)?.ToString() ?? string.Empty,
                Division: row.GetCell(5)?.ToString() ?? string.Empty,
                Pool: row.GetCell(6)?.ToString() ?? string.Empty,
                Team1: row.GetCell(7)?.ToString() ?? string.Empty,
                Team2: row.GetCell(8)?.ToString() ?? string.Empty,
                Referee1: referee1,
                Referee2: referee2,
                Scorer: scorer
            );


            games.Add(gamerecord);
        }
    }
}
else
{
    games.Add(new GameRecord(null, null, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty));
}

Console.WriteLine($"Number of games {games.Count}, creating pdf file.");

PdfDocument document = new PdfDocument();

foreach (var game in games)
{
    PdfPage page = document.AddPage();
    XGraphics gfx = XGraphics.FromPdfPage(page);


    PrintGameRecord(gfx, game);
}

document.Save($"Gamesheet.pdf");
document.Close();


void PrintGameRecord(XGraphics gfx, GameRecord game)
{
    XFont fontBig = new XFont("Verdana", 18, XFontStyle.Bold);
    XFont font = new XFont("Verdana", 9, XFontStyle.Bold);

    var width = (gfx.PageSize.Width - 20) / 4;
    var height = 100;

    double left = 10;
    gfx.DrawRectangle(XPens.Black, new XRect(left, 10, width / 2, height));
    gfx.DrawString($"Tijd:", font, XBrushes.Black, new XRect(left + 5, 15, gfx.PageSize.Width, 20), XStringFormats.CenterLeft);
    gfx.DrawString(game.StartTime.HasValue ? game.StartTime.Value.ToShortTimeString() : string.Empty, fontBig, XBrushes.Black, new XRect(left + 5, 35, gfx.PageSize.Width, 20), XStringFormats.CenterLeft);

    left = (width / 2) + 10;
    gfx.DrawRectangle(XPens.Black, new XRect(left, 10, width / 2, height));

    gfx.DrawString($"Veld:", font, XBrushes.Black, new XRect(left + 5, 15, gfx.PageSize.Width, 20), XStringFormats.CenterLeft);
    gfx.DrawString(game.Field, fontBig, XBrushes.Black, new XRect(left + 5, 35, gfx.PageSize.Width, 20), XStringFormats.CenterLeft);

    double rowHeight = 20;
    left = width + 10;
    double top = 10;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, rowHeight));
    gfx.DrawString("Wedstrijd/Game", font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, rowHeight), XStringFormats.CenterLeft);

    top += rowHeight;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, 2 * rowHeight));
    gfx.DrawString("Scheidsrechters/Ref", font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, 2 * rowHeight), XStringFormats.CenterLeft);

    top += 2 * rowHeight;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, rowHeight));
    gfx.DrawString("Uitslag/Score", font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, rowHeight), XStringFormats.CenterLeft);

    top += rowHeight;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, rowHeight));
    gfx.DrawString("Fairplay", font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, rowHeight), XStringFormats.CenterLeft);

    top = 10;
    left += width;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, rowHeight));
    gfx.DrawString($"A: {game.Team1} - B: {game.Team2}", font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, rowHeight), XStringFormats.CenterLeft);

    top += rowHeight;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, 2 * rowHeight));
    gfx.DrawString(game.Referee1, font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, rowHeight), XStringFormats.CenterLeft);
    gfx.DrawString(game.Referee2, font, XBrushes.Black, new XRect(left + 5, top + 15, gfx.PageSize.Width, rowHeight), XStringFormats.CenterLeft);

    top += 2 * rowHeight;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, rowHeight));
    gfx.DrawString("-", font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, rowHeight), XStringFormats.Center);

    top += rowHeight;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, rowHeight));
    gfx.DrawString("A [5] [4] [3] [2] [1]", font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, rowHeight), XStringFormats.CenterLeft);

    top = 10;
    left += width;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, rowHeight));
    gfx.DrawString($"Poule: {game.Division} {game.Pool}", font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, rowHeight), XStringFormats.CenterLeft);

    top += rowHeight;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, 2 * rowHeight));
    gfx.DrawString(game.Scorer, font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, rowHeight), XStringFormats.CenterLeft);

    top += 2 * rowHeight;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, rowHeight));
    gfx.DrawString("Winnaar:", font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, rowHeight), XStringFormats.CenterLeft);

    top += rowHeight;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, rowHeight));
    gfx.DrawString("B [5] [4] [3] [2] [1]", font, XBrushes.Black, new XRect(left + 5, top, gfx.PageSize.Width, rowHeight), XStringFormats.CenterLeft);

    width = (gfx.PageSize.Width - 20) / 3;

    DrawScoreList(gfx, $"Score Team A: {game.Team1}", "Gebruik 1 regel per score!", 1);
    DrawScoreList(gfx, $"Score Team B: {game.Team2}", "Use one line per score!", 2);

    DrawTeamBox(gfx, $"Team A: {game.Team1}", 1);
    DrawTeamBox(gfx, $"Team B: {game.Team2}", 2);
}

void DrawTeamBox(XGraphics gfx, string team, int teamNumber)
{
    XFont fontBold = new XFont("Verdana", 9, XFontStyle.Bold);
    var width = (gfx.PageSize.Width - 20) / 3;
    var left = 10 + 2 * width;
    double top = teamNumber == 1 ? 110 : 110 + ((12 + 1) * RowHeight) + 25;

    gfx.DrawRectangle(XPens.Black, XBrushes.LightGray, new XRect(left, top, width, 25));
    gfx.DrawString(team, fontBold, XBrushes.Black, new XPoint(left + 5, top + 10));

    top += 25;
    gfx.DrawRectangle(XPens.Black, XBrushes.LightGray, new XRect(left, top, width / 2, RowHeight));
    gfx.DrawString("Speler/Player", fontBold, XBrushes.Black, new XPoint(left + 5, top + 10));
    gfx.DrawRectangle(XPens.Black, XBrushes.LightGray, new XRect(left + width / 2, top, width / 2, RowHeight));
    gfx.DrawString("Fouten/Faults", fontBold, XBrushes.Black, new XPoint(left + width / 2 + 5, top + 10));

    for (var speler = 0; speler < 12; speler++)
    {
        top += RowHeight;
        gfx.DrawRectangle(XPens.Black, new XRect(left, top, width / 2, RowHeight));
        gfx.DrawRectangle(XPens.Black, new XRect(left + width / 2, top, width / 6, RowHeight));
        gfx.DrawRectangle(XPens.Black, new XRect(left + width / 2 + width / 6, top, width / 6, RowHeight));
        gfx.DrawRectangle(XPens.Black, new XRect(left + width / 2 + 2 * width / 6, top, width / 6, RowHeight));
    }
}

void DrawScoreList(XGraphics gfx, string team, string hint, int col)
{
    XFont fontBold = new XFont("Verdana", 9, XFontStyle.Bold);
    XFont font = new XFont("Verdana", 9, XFontStyle.Regular);
    var width = (gfx.PageSize.Width - 20) / 3;
    var left = col == 1 ? 10 : 10 + width;
    double top = 110;

    gfx.DrawRectangle(XPens.Black, XBrushes.LightGray, new XRect(left, top, width, 25));
    gfx.DrawString(team, fontBold, XBrushes.Black, new XPoint(left + 5, top + 10));
    gfx.DrawString("Kleur/Color:", fontBold, XBrushes.Black, new XPoint(left + 5, top + 20));

    top += 25;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width, RowHeight));
    gfx.DrawString(hint, font, XBrushes.Black, new XPoint(left + 5, top + 10));

    top += RowHeight;
    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width / 2, RowHeight));
    gfx.DrawString("1e Helft", font, XBrushes.Black, new XPoint(left + 5, top + 10));

    gfx.DrawRectangle(XPens.Black, new XRect(left + width / 2, top, width / 2, RowHeight));
    gfx.DrawString("2e Helft", font, XBrushes.Black, new XPoint(left + width / 2 + 5, top + 10));

    top += RowHeight;

    gfx.DrawRectangle(XPens.Black, new XRect(left, top, width / 4, RowHeight));
    gfx.DrawString("Speler", font, XBrushes.Black, new XPoint(left + 5, top + 10));

    gfx.DrawRectangle(XPens.Black, new XRect(left + width / 4, top, width / 4, RowHeight));
    gfx.DrawString("Punt", font, XBrushes.Black, new XPoint(left + width / 4 + 5, top + 10));

    gfx.DrawRectangle(XPens.Black, new XRect(left + (2 * width / 4), top, width / 4, RowHeight));
    gfx.DrawString("Speler", font, XBrushes.Black, new XPoint(left + (2 * width / 4) + 5, top + 10));

    gfx.DrawRectangle(XPens.Black, new XRect(left + (3 * width / 4), top, width / 4, RowHeight));
    gfx.DrawString("Punt", font, XBrushes.Black, new XPoint(left + (3 * width / 4) + 5, top + 10));

    for (var row = 0; row < 32; row++)
    {
        top += RowHeight;
        gfx.DrawRectangle(XPens.Black, new XRect(left, top, width / 4, RowHeight));

        gfx.DrawRectangle(XPens.Black, new XRect(left + width / 4, top, width / 4, RowHeight));

        gfx.DrawRectangle(XPens.Black, new XRect(left + (2 * width / 4), top, width / 4, RowHeight));

        gfx.DrawRectangle(XPens.Black, new XRect(left + (3 * width / 4), top, width / 4, RowHeight));
    }
}

public record GameRecord(DateTime? StartTime, DateTime? EndTime, string Day, string Field, string Phase, string Division, string Pool, string Team1, string Team2, string Referee1, string Referee2, string Scorer);

