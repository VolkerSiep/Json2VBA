# Change log

## V0.2 (main branch - not yet released)
- Consider "," as decimal separator when parsing numbers
- Replace `\t` and `\n` from json strings with `vbCrLf` and `vbTab`
- Declare `current_token` as `long` to prevent overflow for large json data sets

## V0.1 (Released May 18th 2023)
- initial release

Known issues:
 - Didn't care yet about users with "," as decimal separator.
