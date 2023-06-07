
# QB Checker

QB Checker is for use on the Infinite Leaftide Asheron's Call server. It will compare your QB stamps to a list of known stamps and tell you if you have collected them or not.

## Getting started

1. Either download the [latest releases](https://github.com/Scoboose/QBChecker/releases) or build from source

2. Download the [Master QB List](https://github.com/Scoboose/QBChecker/blob/master/MasterQBList.csv)

2. Export your full `/qb list` from in-game
* Disable chat time stamps
* /log qb
* /qb list
* /log

3. Place both the MasterQBList and your QB export into the QB checker folder

4. Run QBChecker.exe via Powershell or CMD
* Right click on empty space in the qbchecker foler and click `Open in Terminal`
* Type `.\QBChecker.exe .\MasterQBList.csv .\qb.txt`

5. Open My_QB_List.xlsx

### Notes

If you do not have an application on your computer that can open excel documents you can use [Google Sheets](https://www.google.com/sheets)

### Please help!
If you have stamps not on the master QB list and know how you receaved them please update the MasterQBList.csv and do make a pull request. Or just contact Pokey ❤️
