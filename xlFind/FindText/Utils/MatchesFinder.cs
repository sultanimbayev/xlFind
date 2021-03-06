﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using IEIT.Reports.Export.Helpers.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using xlFind.Models.Interfaces;
using xlFind.Models.Replace;
using xlFind.Models.Search;

namespace xlFind.Utils
{
    public class MatchesFinder
    {

        public static IEnumerable<FileMatch> Find(string dir, string searchTxt, bool useRegex)
        {
            return new MatchesFinder().FindMatches(dir, searchTxt, useRegex);
        }

        public IEnumerable<FileMatch> FindMatches(string dir, string searchTxt, bool useRegex)
        {
            var filepathsList = FilesFinder.FindFrom(dir).WithExtension("xlsx");
            if (!useRegex) { searchTxt = Regex.Escape(searchTxt); }
            Regex rgx = new Regex(searchTxt, RegexOptions.IgnoreCase);
            var result = new List<FileMatch>();
            foreach (var filepath in filepathsList)
            {
                try
                {
                    var worksheetMatches = FindMatchesInFile(filepath, rgx).ToList();
                    if(worksheetMatches.Count > 0)
                    {
                        result.Add(new FileMatch()
                        {
                            Filepath = filepath,
                            WorksheetMatches = worksheetMatches
                        });
                    }

                }
                catch (Exception e)
                {
                    Console.WriteLine($"Error on reading file: {filepath}");
                    Console.WriteLine(e.Message);
#if DEBUG
                    Console.WriteLine(e.StackTrace);
#endif
                }
            }
            return result;
        }

        public IEnumerable<WorksheetMatch> FindMatchesInFile(string filepath, Regex rgx)
        {
            var doc = SpreadsheetDocument.Open(filepath, false);
            foreach (var ws in GetWorksheets(doc))
            {
                var worksheetMatch = GetMathcesInWorksheet(ws, rgx);
                if(worksheetMatch != null)
                {
                    yield return worksheetMatch;
                }
            }
            doc.Close();
        }

        public WorksheetMatch GetMathcesInWorksheet(Worksheet ws, Regex rgx)
        {
            var cells = ws.FindCells(rgx);
            if (cells != null && cells.Count() > 0)
            {
                var worksheetMatch = new WorksheetMatch() { SheetName = ws.GetName() };
                foreach (var cell in cells)
                {
                    var cellMatch = new CellMatch()
                    {
                        CellAddress = cell.CellReference?.Value,
                        Value = cell.GetValue()
                    };
                    worksheetMatch.CellMatches.Add(cellMatch);
                }
                return worksheetMatch;
            }
            return null;
        }

        public static IEnumerable<FileMatch> Replace(IEnumerable<FileMatch> fileMatches, string searchTxt, string replacement, bool useRegex)
        {
            if (!useRegex) { searchTxt = Regex.Escape(searchTxt); }
            Regex rgx = new Regex(searchTxt, RegexOptions.IgnoreCase);
            var replaces = fileMatches.Select(fm => fm.Clone()).ToList();

            foreach (var fileMatch in fileMatches)
            {
                var doc = SpreadsheetDocument.Open(fileMatch.Filepath, true);

                foreach (var worksheetMatch in fileMatch.WorksheetMatches)
                {
                    var ws = doc.GetWorksheet(worksheetMatch.SheetName);
                    var replacedWorksheetMatch = replaces
                            .FirstOrDefault(fm => fm.Filepath.Equals(fileMatch.Filepath))?
                            .WorksheetMatches?
                            .FirstOrDefault(wm => wm.SheetName.Equals(worksheetMatch.SheetName));

                    if (replacedWorksheetMatch == null) { continue; }

                    //var cellMathces = 
                    replacedWorksheetMatch.CellMatches = new List<ICellMatch>();

                    foreach (var cellMatch in worksheetMatch.CellMatches)
                    {
                        var cell = ws.GetCell(cellMatch.CellAddress);
                        var cellOldValue = cell.GetValue();

                        cell = RegexReplaceTextInCell(cell, rgx, replacement);
                        var cellNewValue = cell.GetValue();

                        var replacedValue = new ReplacedCellMatch()
                        {
                            CellAddress = cell.CellReference?.Value,
                            OldValue = cellOldValue,
                            Value = cellNewValue
                        };


                        replacedWorksheetMatch.CellMatches.Add(replacedValue);

                    }
                }
                doc.SaveAndClose();
            }
            return replaces;
        }

        static IEnumerable<Cell> RegexRepleceTextInCells(IEnumerable<Cell> cells, Regex rgx, string replacement)
        {
            foreach (var cell in cells)
            {
                var newCell = RegexReplaceTextInCell(cell, rgx, replacement);
                yield return newCell;
            }
        }
        static Cell RegexReplaceTextInCell(Cell cell, Regex rgx, string replacement)
        {
            var oldValue = cell.GetValue();
            var newValue = rgx.Replace(oldValue, replacement);
            var ws = cell.GetWorksheet();
            var cellAddr = cell.CellReference.Value;
            cell.WriteText(newValue);
            return ws.GetCell(cellAddr);
        }
        static IEnumerable<Worksheet> GetWorksheets(SpreadsheetDocument doc)
        {
            var sheets = doc.WorkbookPart.Workbook.Sheets;
            foreach (Sheet sheet in sheets)
            {
                var wsPart = doc.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;
                var ws = wsPart?.Worksheet;
                if (ws == null)
                {
                    Console.WriteLine($"Attemption of getting worksheet '{sheet.Name}' failed. It is not a worksheet!");
                    continue;
                }
                yield return ws;
            }
        }
    }
}
