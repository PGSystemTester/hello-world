# Find Replace REGEX in NOTEPAD++

Removes blank line between two lines starting with //

**Before**
`````
// bla blah note

/// second line
`````

**After**
`````
// bla blah note
/// second line
`````

- FIND:`(^\/\/[\s\S]*?)(?:\r?\n){2}(?=\/\/)`
- REPLACE: `$1\n`



# Other Regex Cheat Sheets
- [https://cheatography.com/davechild/cheat-sheets/regular-expressions/](https://cheatography.com/davechild/cheat-sheets/regular-expressions/)
