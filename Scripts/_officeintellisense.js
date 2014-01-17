// Required to correctly initalize IntelliSense in Visual Studio

var Office = new function () {
    this._appContext = 7;
};

// 1 Excel appContext = 001
// 2 Word appContext = 010
// 3 Word + Excel appContext = 011
// 4 Project appContext = 100
// 5 Project + Excel appContext = 101
// 6 Project + Word appContext = 110
// 7 Project + Word + Excel appContext = 111
// 8 Outlook appContext = 1000
// 16 PowerPoint appContext
// 23 PowerPoint + Word + Excel appContext = 101111
