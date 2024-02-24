# GetDoW

This script determines the day of the week of a specific date. The date can be provided via command line or stdin. The day of the week can be displayed in numerical form, as german or english abbreviation or with its full german or english name.

The code is based on the algorithm of german mathematician Christian Zeller named [_Zeller's congruence_](https://en.wikipedia.org/wiki/Zeller%27s_congruence). It uses the Gregorian calendar variant for calculating the day of the week.


## Translation of the script's help

```Z
Syntax: GetDoW Date|/i [/m:mask] [/n|/gs|/gl|/es|/el]

Parameter: 1. Date
              2-digits day and month with leading zero, year 4-digits
              Mandatory parameter (except parameter /i is provided)
              The delimiter character has to be the same like in the mask.

           2. /m:mask
              Mask that defines where day, month and year are located in
              provided date (for example YYYY-MM-DD). Supported delimiters
              are -./
              Default DD.MM.YYYY (german date format)

           3. Type of output
                /n  --> Numeric (Monday=1...Sunday=7)
                /gs --> German abbreviation of day of week
                /gl --> Full german name of day of week
                /es --> English abbreviation of day of week
                /el --> Full english name of day of week
              Default /n

           4. /i
              If this switch is provided but no date, the date will be
              read from STDIN.
              So the following is possible:
                ECHO %DATE% | getdow /i /gl (german Windows)
                  or
                ECHO %DATE% | getdow /i /el /m:YYYY-MM-DD (US Windows)
                  or
                ECHO. | SET /P=Heute ist & ECHO %DATE% | getdow /i /gl
                (german Windows)
                  or
                ECHO. | SET /P=Today is & ECHO %DATE% | getdow /i /el /m:YYYY-MM-DD
                (US Windows)

The order of the parameters is arbitrary.
In case of an error ERRORLEVEL will be set to 1, otherwise to 0.
Algorithm according to Christian Zeller (Zeller's congruence).
```
