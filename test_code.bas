Attribute VB_Name = "Module1"
Sub Stock_Check()
    '   Variable declarations
    Dim WS As Worksheet

    Dim current_row As Long
    Dim start_row As Long
    Dim end_row As Long
    Dim print_row As Long

    Dim greatest_increase As Double
    Dim greatest_decrease As Double
    Dim greatest_volume As Double

    Dim yearly_change As Double
    Dim percentage_change As Double
    Dim cumulative_total As Double
    Dim current_ticker As String

    '   Iterate through worksheets in the workbook
    '----------------------------------------------------------------------
    For Each WS In ThisWorkbook.Worksheets

        WS.Activate

        '   Set additional column headers
        Range("I1").Value = "Ticker"
        Range("J1").Value = "Yearly Change"
        Range("K1").Value = "Percentage Change"
        Range("L1").Value = "Total Stock Volume"
        Range("O2").Value = "Greatest % Increase"
        Range("O3").Value = "Greatest % Decrease"
        Range("O4").Value = "Greatest Total Volume"
        Range("P1").Value = "Ticker"
        Range("Q1").Value = "Value"

        '   Set starting values for key variables
        current_row = 2
        start_row = 2
        print_row = 2
        current_ticker = Cells(current_row, 1).Value

        '   Move sequentially down rows until end is reached (first blank cell)
        '----------------------------------------------------------------------
        Do While IsEmpty(Cells(current_row, 1).Value) <> True

            '   Check if new row is a new ticker
            '   If so, summate and display results for previous ticker to the table
            If Cells(current_row, 1).Value <> current_ticker Then
                end_row = current_row - 1

                '   if start row is still 0 the stock never traded (so yearly change is 0)
                If start_row = 0 Then
                    yearly_change = 0
                Else
                    yearly_change = (Cells(end_row, 6).Value) - (Cells(start_row, 3).Value)
                End If

                '   percentage change = yearly change / opening price
                If yearly_change <> 0 Then
                    percentage_change = Round((yearly_change / Cells(start_row, 3).Value), 4)
                Else
                    percentage_change = 0
                End If

                '   Display results in table with conditional formatting
                Call Display_Results(print_row, current_ticker, yearly_change, percentage_change, cumulative_total)

                '   Reset variable values for the new ticker
                '   Check if the start value for the new ticker is 0 (i.e. didn't start trading on Jan 1)
                If Cells(current_row, 3).Value <> 0 Then
                    start_row = current_row
                Else
                    start_row = 0  '   if so, set to placeholder value of 0 until it's identified that trading has started
                End If
                cumulative_total = Cells(current_row, 7).Value
                print_row = print_row + 1
                current_ticker = Cells(current_row, 1).Value

            '   If new row is not a new ticker just add on the days traded volume
            Else
                '   check if trading has started for new stocks
                If (Cells(current_row, 3).Value <> 0) And (start_row = 0) Then
                    start_row = current_row
                End If
                cumulative_total = cumulative_total + Cells(current_row, 7).Value

            End If

            current_row = current_row + 1

        Loop

        '   Call display function again, otherwise final ticker is missed
        Call Display_Results(print_row, current_ticker, yearly_change, percentage_change, cumulative_total)


        '   Find the greatest values from the results list
        '----------------------------------------------------------------------
        current_row = 2

        Range("P2").Value = Cells(current_row, 9).Value
        greatest_increase = Cells(current_row, 11).Value
        greatest_decrease = Cells(current_row, 11).Value
        greatest_volume = Cells(current_row, 12).Value

        Do While IsEmpty(Cells(current_row, 9).Value) <> True
            If Cells(current_row, 11).Value >= greatest_increase Then
                greatest_increase = Cells(current_row, 11).Value
                Range("P2").Value = Cells(current_row, 9).Value '   ticker
                Range("Q2").Value = greatest_increase           '   value
                Range("Q2").NumberFormat = "0.00%"
            ElseIf Cells(current_row, 11).Value < greatest_decrease Then
                greatest_decrease = Cells(current_row, 11).Value
                Range("P3").Value = Cells(current_row, 9).Value '   ticker
                Range("Q3").Value = greatest_decrease           '   value
                Range("Q3").NumberFormat = "0.00%"
            End If

            If (Cells(current_row, 12).Value > greatest_volume) Then
                greatest_volume = Cells(current_row, 12).Value
                Range("P4").Value = Cells(current_row, 9).Value '   ticker
                Range("Q4").Value = greatest_volume             '   value
            End If

            current_row = current_row + 1

        Loop

        '   Additional formatting
        Range("I:Q").Columns.AutoFit
        ActiveWindow.ScrollRow = 1

    Next WS

End Sub

Sub Display_Results(print_row, current_ticker, yearly_change, percentage_change, cumulative_total)

    '   Outputs
    Cells(print_row, 9).Value = current_ticker
    Cells(print_row, 10).Value = yearly_change
    Cells(print_row, 11).Value = percentage_change
    Cells(print_row, 11).NumberFormat = "0.00%"
    Cells(print_row, 12).Value = cumulative_total

    '   Conditional formatting
    If yearly_change < 0 Then
        Cells(print_row, 10).Interior.Color = RGB(255, 0, 0)
    Else
        Cells(print_row, 10).Interior.Color = RGB(0, 255, 0)
    End If

End Sub

Sub Extra_Function()

  'test function that does nothing

End Sub


# Dependencies and Setup
import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import plotly.express as px

import gmaps
import requests
import json


from scipy.stats import linregress

# Import API key
from api_keys import weather_api_key
from api_keys import g_key
# Import World Happiness Report Data 2021
raw_happiness_df = pd.read_csv("Data/world-happiness-report-2021.csv")

# Rename columns
raw_happiness_df = raw_happiness_df.rename(columns={'Country name': 'Country', 
                                                'Regional indicator': 'Region',
                                                'Ladder score': 'Happiness Score',
                                                'Social support': 'Social Support',                                                    
                                                'Logged GDP per capita': 'GDP per Capita',
                                                'Healthy life expectancy': 'Life Expectancy',
                                                'Freedom to make life choices': 'Freedom',
                                                'Perceptions of corruption': 'Corruption'})

# Drop columns not needed
happiness_df = raw_happiness_df.drop(columns=['Standard error of ladder score', 
                                              'upperwhisker', 
                                              'lowerwhisker',
                                              'Ladder score in Dystopia',
                                              'Explained by: Log GDP per capita',
                                              'Explained by: Social support',
                                              'Explained by: Healthy life expectancy',
                                              'Explained by: Freedom to make life choices',
                                              'Explained by: Generosity',
                                              'Explained by: Perceptions of corruption',
                                              'Dystopia + residual'])
# Show preview of DataFrame
happiness_df
