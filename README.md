# X-PLAN

X-PLAN is Excel 2010 template which will help you to quicly generate simple visual schedules.

![demo 01](https://github.com/JaroslavPlavec/x-plan/blob/media/demo01.gif)

## Description

1. X-PLAN is weekly schedule - each column is one week

![explanation 01](https://github.com/JaroslavPlavec/x-plan/blob/media/explanation01.png)

### CALENDAR WEEK

```
=TEXT(WEEKNUM(E1;21);"00")

```

Calendar weeks are in format suitable for Europe. Return type is "21", which means week 1 is the week containing the first Thursday of the year, following ISO 8601.


### MONDAYS

This is just to show what is the date of Monday in relevant week. I.e. in CW02/2019, the date of Monday is 07.01.2019.



### HIGHLIGHTING OF WEEK COLUMN (after mouse click)

This is done by combination of Conditional Formating and VBA.

1. In VBA, there is this code connected to the SHEET:
```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
Range("B1").Calculate
End Sub
```
2. In cell B1, there is this formula:

```
=CELL("col")
```

3. In Conditional Formating, related column is defined like this:


![explanation 02](https://github.com/JaroslavPlavec/x-plan/blob/media/explanation02.png)


