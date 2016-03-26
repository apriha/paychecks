Option Explicit

' *****************************************************************************
' paychecks
'
' VBA code to generate an Excel workbook with dynamic pie charts for
' tracking paychecks
'
' Usage: Import into Excel's Visual Basic Editor and run main
'
' Copyright (C) 2016, Andrew Riha
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
' *****************************************************************************

Public Sub main()
' Generates a workbook with dynamic pie charts for tracking paychecks
' Note: May need to close all workbooks before running

    ' BEGIN EDIT CONSTANTS ----------------------------------------------------
    
    ' column types and quantities
    Const EARNINGS_COLUMNS As Integer = 4    ' 1 or greater
    Const BEFORE_TAX_COLUMNS As Integer = 1  ' 0 or greater
    Const AFTER_TAX_COLUMNS As Integer = 0   ' 0 or greater
    Const TAX_COLUMNS As Integer = 3         ' 0 or greater
    
    ' initial number of paychecks / rows for paychecks
    Const PAYCHECK_ROWS As Integer = 26
    
    ' END EDIT CONSTANTS ------------------------------------------------------
    
    Dim last_column As Integer  ' last column of data
        
    Application.ScreenUpdating = False
        
    'Create workbook
    Workbooks.Add
    
    last_column = add_columns(EARNINGS_COLUMNS, BEFORE_TAX_COLUMNS, AFTER_TAX_COLUMNS, TAX_COLUMNS)
    Call add_paycheck_rows(PAYCHECK_ROWS)
    Call add_grand_total_formulas(PAYCHECK_ROWS, last_column)
    Call add_stuff_for_dynamic_pie_charts(PAYCHECK_ROWS, EARNINGS_COLUMNS, last_column)
    Call add_pie_charts

    Application.ScreenUpdating = True
End Sub

Private Function add_columns(EARNINGS_COLUMNS As Integer, BEFORE_TAX_COLUMNS As Integer, AFTER_TAX_COLUMNS As Integer, TAX_COLUMNS As Integer) As Integer
    ' Add paycheck column titles to workbook; add_column_group does the heavy lifting for each group
    '
    '   Inputs:
    '       EARNINGS_COLUMNS: Integer, number of columns for income
    '       BEFORE_TAX_COLUMNS: Integer, number of columns for before tax contributions
    '       AFTER_TAX_COLUMNS: Integer, number of columns for after tax contributions
    '       TAX_COLUMNS: Integer, number of columns for taxes
    '   Output:
    '       Integer, column number of Net Pay column (last column of data)
    
    Dim earnings_total_cell As Range
    Dim before_tax_total_cell As Range
    Dim after_tax_total_cell As Range
    Dim tax_total_cell As Range
    Dim net_pay_cell As Range

    Range("A2").Value = "Paycheck Number"
    Range("B2").Value = "Date"
    
    Set earnings_total_cell = add_column_group(Range("B2"), "Earnings", EARNINGS_COLUMNS)
    Set before_tax_total_cell = add_column_group(earnings_total_cell, "Before Tax", BEFORE_TAX_COLUMNS)
    Set after_tax_total_cell = add_column_group(before_tax_total_cell, "After Tax", AFTER_TAX_COLUMNS)
    Set tax_total_cell = add_column_group(after_tax_total_cell, "Tax", TAX_COLUMNS)
    Set net_pay_cell = add_column_group(tax_total_cell, "Net Pay", 1)
    
    Call add_net_pay_total_formula(earnings_total_cell.Offset(1, 0), before_tax_total_cell.Offset(1, 0), after_tax_total_cell.Offset(1, 0), tax_total_cell.Offset(1, 0), net_pay_cell.Offset(1, 0))
    
    Rows("1:2").Font.Bold = True
    
    add_columns = net_pay_cell.Column
End Function

Private Function add_column_group(previous_total_cell As Range, base_title As String, columns As Integer) As Range
    ' Add paycheck columns to workbook, group columns of same type, and add total formula for groups
    '
    '   Inputs:
    '       previous_total: Range, cell of previous "Total" value
    '       base_title: String, title used for columns and group of columns
    '       columns: Integer, columns to add for base_title group
    '   Output:
    '       Range, cell of group's last header cell (i.e., "Total" cell or "Net Pay" cell)
    
    Dim i As Integer
    Dim address_group_start As String
    Dim address_group_end As String
    
    previous_total_cell.Activate
    
    If base_title = "Net Pay" Then
        ActiveCell.Offset(0, 1).Activate
        ActiveCell.Value = "Net Pay"
        Call group_titles("Net Pay", ActiveCell.Offset(-1, 0).Address, ActiveCell.Offset(-1, 0).Address)
    ElseIf columns > 0 Then
        For i = 1 To columns
            ActiveCell.Offset(0, i).Value = base_title & " " & i
            
            If i = 1 Then
                'Offset by one row above (currently empty, will be used to group columns of same type)
                address_group_start = ActiveCell.Offset(-1, i).Address
            End If
        Next i
            
        ActiveCell.Offset(0, i).Activate
        ActiveCell.Value = "Total"
        address_group_end = ActiveCell.Offset(-1, 0).Address
        
        Call group_titles(base_title, address_group_start, address_group_end)
        Call add_group_total_formula(Range(address_group_start).Offset(2, 0), Range(address_group_end).Offset(2, 0))
    End If
    
    Set add_column_group = ActiveCell
End Function

Private Sub group_titles(base_title As String, address_group_start As String, address_group_end As String)
    ' Group titles and format groups
    
    Dim group_range As Range

    Set group_range = Range(address_group_start & ":" & address_group_end)
    
    If address_group_start <> address_group_end Then
        group_range.Merge
    End If
    
    Select Case base_title
        Case "Earnings"
            group_range.Value = "Earnings"
            group_range.Interior.Color = RGB(153, 204, 255)  ' blue 99CCFF
        Case "Before Tax"
            group_range.Value = "Before Tax Deductions"
            group_range.Interior.Color = RGB(204, 255, 204)  ' green CCFFCC
        Case "After Tax"
            group_range.Value = "After Tax Deductions"
            group_range.Interior.Color = RGB(204, 255, 255)  ' lightblue CCFFFF
        Case "Tax"
            group_range.Value = "Taxes"
            group_range.Interior.Color = RGB(255, 204, 153)  ' tan FFCC99
        Case "Net Pay"
            group_range.Value = "Net Pay"
            group_range.Interior.Color = RGB(252, 203, 44)  ' gold FFCC00
    End Select
End Sub

Private Sub add_group_total_formula(start_cell As Range, end_cell As Range)
    ' Add total formula for the first row of a group and format
    
    end_cell.Formula = "=SUM(" & start_cell.Address(False, False) & ":" & end_cell.Offset(0, -1).Address(False, False) & ")"
    Range(start_cell, end_cell).NumberFormat = "$#,##0.00"
    Call format_calculated_cell(end_cell)
End Sub

Private Sub add_net_pay_total_formula(source_total_cell As Range, before_tax_total_cell As Range, after_tax_total_cell As Range, tax_total_cell As Range, net_pay_cell As Range)
    ' Add net pay formula; derive from all address of group "Total" cells
    
    Dim net_pay_formula As String

    net_pay_formula = "=(" & source_total_cell.Address(False, False)
    
    ' If a group does not have any columns, the address will be the same as the previous group
    If source_total_cell.Address <> before_tax_total_cell.Address Then
        net_pay_formula = net_pay_formula & "-" & before_tax_total_cell.Address(False, False)
    End If
    
    If before_tax_total_cell.Address <> after_tax_total_cell.Address Then
        net_pay_formula = net_pay_formula & "-" & after_tax_total_cell.Address(False, False)
    End If
    
    If after_tax_total_cell.Address <> tax_total_cell.Address Then
        net_pay_formula = net_pay_formula & "-" & tax_total_cell.Address(False, False)
    End If
    
    net_pay_formula = net_pay_formula & ")"
    
    net_pay_cell.Formula = net_pay_formula
    
    Call format_calculated_cell(net_pay_cell)
End Sub

Private Sub format_calculated_cell(calculated_cell As Range)
    ' Format a calculated cell with gray background and regular borders

    With calculated_cell.Interior
        .Color = RGB(242, 242, 242)
    End With
    With calculated_cell.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(206, 206, 206)
    End With
    With calculated_cell.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(206, 206, 206)
    End With
    With calculated_cell.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(206, 206, 206)
    End With
    With calculated_cell.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(206, 206, 206)
    End With
End Sub

Private Sub add_paycheck_rows(PAYCHECK_ROWS As Integer)
    ' Adds rows for paychecks
    
    Dim i As Integer

    Range("A3").Activate

    For i = 1 To PAYCHECK_ROWS
        ActiveCell.Offset(i - 1, 0).Value = i
    Next i
End Sub

Private Sub add_grand_total_formulas(PAYCHECK_ROWS As Integer, last_column As Integer)
    ' Add total formulas for columns to a grand total row
    
    Dim grand_total_row As Integer
    
    grand_total_row = PAYCHECK_ROWS + 3 ' 2 rows for headers + 1 for grand_total_row

    Range("A" & grand_total_row).Activate
    ActiveCell.Value = "Grand Total"
    ActiveCell.Font.Bold = True
    Call format_calculated_cell(ActiveCell.Offset(0, 2))
    
    ' Fill down all formatting and total formulas
    Range("C3:" & ActiveCell.Offset(-1, last_column - 1).Address).FillDown
        
    With Range(ActiveCell, ActiveCell.Offset(0, last_column - 1)).Borders(xlEdgeTop)
        .LineStyle = xlDouble
        .Weight = xlThick
        .ColorIndex = 1
    End With
    
    ActiveCell.Offset(0, 2).Activate
    ActiveCell.FormulaR1C1 = "=SUM(R[-" & PAYCHECK_ROWS & "]C:R[-1]C)"
        
    Range(ActiveCell, ActiveCell.Offset(0, last_column - 3)).FillRight
End Sub

Private Sub add_stuff_for_dynamic_pie_charts(PAYCHECK_ROWS As Integer, EARNINGS_COLUMNS As Integer, last_column As Integer)
    ' Dynamic pie charts require some temp data, named ranges, and deep magic
    
    Dim pie_chart_data_start_row As Integer
    Dim earnings_columns_total As Integer
    Dim destination_columns_total As Integer
    Dim row_titles_top_cell As Range
    Dim source_top_left_cell As Range
    Dim destination_top_left_cell As Range
    
    pie_chart_data_start_row = PAYCHECK_ROWS + 4 ' paycheck rows + 2 rows for headers + 1 row for grand totals + 1 row for pie_chart_data_start_row
    earnings_columns_total = EARNINGS_COLUMNS + 1 ' earnings columns + 1 column for earnings "Total" column
    destination_columns_total = last_column - earnings_columns_total - 2 ' total columns - total earnings columns - 2 columns for paycheck number and paycheck date
    
    Set row_titles_top_cell = Range("A" & pie_chart_data_start_row)
    
    ' Title formula rows
    row_titles_top_cell.Offset(0, 0).Value = "Source Temp Data"
    row_titles_top_cell.Offset(1, 0).Value = "Source Temp Data"
    row_titles_top_cell.Offset(2, 0).Value = "Source Data"
    row_titles_top_cell.Offset(3, 0).Value = "Source Labels"
    row_titles_top_cell.Offset(4, 0).Value = "Destination Temp Data"
    row_titles_top_cell.Offset(5, 0).Value = "Destination Temp Data"
    row_titles_top_cell.Offset(6, 0).Value = "Destination Data"
    row_titles_top_cell.Offset(7, 0).Value = "Destination Labels"
    
    ' Source temp data formulas
    Set source_top_left_cell = Range("A" & pie_chart_data_start_row).Offset(0, 2)
    source_top_left_cell.Offset(0, 0).FormulaR1C1 = "=IF(R[-" & PAYCHECK_ROWS + 2 & "]C=""Total"", 0, R[-1]C+(COLUMNS(R[-1]C3:R[-1]C)/1000*(R[-1]C<>0)))"
    source_top_left_cell.Offset(1, 0).FormulaR1C1 = "=MATCH(SMALL(R[-1],COUNTIF(R[-1],0)+COLUMNS(R[-1]C3:R[-1]C)),R[-1],0)"
    
    ' Source pie chart data / label formulas
    source_top_left_cell.Offset(2, 0).FormulaR1C1 = "=OFFSET(R" & PAYCHECK_ROWS + 2 & "C2,1,MATCH(INDEX(R[-2],SMALL(OFFSET(R[-1]C3,0,0,1,COUNTIF(R[-2],"">0"")),COLUMNS(R[-2]C3:R[-2]C))),R[-2]C3:R[-2]C" & EARNINGS_COLUMNS + 3 & ",0),1,1)"
    source_top_left_cell.Offset(3, 0).FormulaR1C1 = "=OFFSET(R[-" & PAYCHECK_ROWS + 5 & "]C2,0,MATCH(INDEX(R[-3],1,SMALL(OFFSET(R[-2]C3,0,0,1,COUNTIF(R[-3],"">0"")),COLUMNS(R[-3]C3:R[-3]C))),R[-3]C3:R[-3]C" & EARNINGS_COLUMNS + 3 & ",0),1,1)"
    
    Call fill_dynamic_pie_chart_formulas(source_top_left_cell, EARNINGS_COLUMNS)
    
    ' Destination temp data formulas
    Set destination_top_left_cell = Range("A" & pie_chart_data_start_row).Offset(4, 2 + earnings_columns_total)
    destination_top_left_cell.Offset(0, 0).Activate
    ActiveCell.FormulaR1C1 = "=IF(R[-" & PAYCHECK_ROWS + 6 & "]C=""Total"", 0, R[-5]C+(COLUMNS(R[-5]C" & ActiveCell.Column & ":R[-5]C)/1000*(R[-5]C<>0)))"
    
    destination_top_left_cell.Offset(1, 0).Activate
    ActiveCell.FormulaR1C1 = "=MATCH(SMALL(R[-1],COUNTIF(R[-1],0)+COLUMNS(R[-1]C" & ActiveCell.Column & ":R[-1]C)),R[-1],0)"
    
    ' Destination pie chart data / label formulas
    destination_top_left_cell.Offset(2, 0).Activate
    ActiveCell.FormulaR1C1 = "=OFFSET(R" & PAYCHECK_ROWS + 2 & "C" & ActiveCell.Column - 1 & ",1,MATCH(INDEX(R[-2],SMALL(OFFSET(R[-1]C" & ActiveCell.Column & ",0,0,1,COUNTIF(R[-2],"">0"")),COLUMNS(R[-2]C:R[-2]C" & ActiveCell.Column & "))),R[-2]C" & ActiveCell.Column & ":R[-2]C" & ActiveCell.Column + destination_columns_total - 1 & ",0),1,1)"
    
    destination_top_left_cell.Offset(3, 0).Activate
    ActiveCell.FormulaR1C1 = "=OFFSET(R[-" & PAYCHECK_ROWS + 9 & "]C" & ActiveCell.Column - 1 & ",0,MATCH(INDEX(R[-3],1,SMALL(OFFSET(R[-2]C" & ActiveCell.Column & ",0,0,1,COUNTIF(R[-3],"">0"")),COLUMNS(R[-3]C" & ActiveCell.Column & ":R[-3]C))),R[-3]C" & ActiveCell.Column & ":R[-3]C" & ActiveCell.Column + destination_columns_total - 1 & ",0),1,1)"
    
    Call fill_dynamic_pie_chart_formulas(destination_top_left_cell, destination_columns_total - 1)

    ' Name ranges for dynamic pie charts
    source_top_left_cell.Offset(2, 0).Activate
    ActiveWorkbook.Names.Add Name:="SourcePieData", RefersToR1C1:="=OFFSET(Sheet1!R" & ActiveCell.Row & "C3,0,0,1,MAX(1,COUNT(Sheet1!R" & ActiveCell.Row & "C3:R" & ActiveCell.Row & "C" & EARNINGS_COLUMNS + 3 & ")))"
    ActiveWorkbook.Names.Add Name:="SourcePieLabels", RefersToR1C1:="=OFFSET(SourcePieData,1,0)"
    
    source_top_left_cell.Offset(6, 0).Activate
    ActiveWorkbook.Names.Add Name:="DestinationPieData", RefersToR1C1:="=OFFSET(Sheet1!R" & ActiveCell.Row & "C" & ActiveCell.Column + earnings_columns_total & ",0,0,1,MAX(1,COUNT(Sheet1!R" & ActiveCell.Row & "C" & ActiveCell.Column + earnings_columns_total & ":R" & ActiveCell.Row & "C" & ActiveCell.Column + earnings_columns_total + destination_columns_total & ")))"
    ActiveWorkbook.Names.Add Name:="DestinationPieLabels", RefersToR1C1:="=OFFSET(DestinationPieData,1,0)"
    
    Rows(pie_chart_data_start_row & ":" & pie_chart_data_start_row + 7).EntireRow.Hidden = True
End Sub

Private Sub fill_dynamic_pie_chart_formulas(top_left_cell As Range, columns As Integer)
    ' Fill formulas for dynamic pie charts
    
    If columns <> 0 Then
        Range(top_left_cell, top_left_cell.Offset(3, columns)).FillRight
    End If
End Sub

Private Sub add_pie_charts()
    ' Add dynamic pie charts - one for Pay Source, one for Pay Destination
    
    Dim workbook_name As String
    
    workbook_name = ActiveWorkbook.Name
    
    ActiveSheet.Shapes.AddChart.Select
    Call format_pie_chart(workbook_name, ActiveChart, "Pay Source", "SourcePieData", "SourcePieLabels")
    
    ActiveSheet.Shapes.AddChart.Select
    Call format_pie_chart(workbook_name, ActiveChart, "Pay Destination", "DestinationPieData", "DestinationPieLabels")
End Sub

Private Sub format_pie_chart(workbook_name As String, active_chart As Chart, title As String, data As String, labels As String)
    ' Jazz up the pie chart with data, labels, and formatting
    
    active_chart.ChartType = xlPie
    active_chart.HasLegend = False
    active_chart.PlotVisibleOnly = False
    active_chart.SeriesCollection.NewSeries
    active_chart.SeriesCollection(1).Name = title
    active_chart.SeriesCollection(1).Values = (workbook_name & "!" & data)
    active_chart.SeriesCollection(1).XValues = (workbook_name & "!" & labels)
    active_chart.SeriesCollection(1).HasDataLabels = True
    active_chart.SeriesCollection(1).DataLabels.ShowValue = False
    active_chart.SeriesCollection(1).DataLabels.ShowCategoryName = True
    active_chart.SeriesCollection(1).DataLabels.ShowPercentage = True
    active_chart.SeriesCollection(1).DataLabels.NumberFormat = "0.00%"
    active_chart.SeriesCollection(1).HasLeaderLines = True
End Sub
