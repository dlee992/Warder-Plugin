# Warder-Plugin Development

## Run code on Mac OS

    npm run dev-server //start the local web server if developing on Mac
    npm start          //test on the desktop
    npm run start:web  //test on a browser

Use online onedrive Excel, insert your add-in manifest.xml file, then can automatically load and test on browser

## JS Syntax (for quick, elegant coding)

可选链

    let user = {}
    alert( user?.address?.street)

## Debug

Office-js 使用 TS 的编译链，但我仍然使用许多 JS 的第三方库，很多库没有类型声明文件（dts），那么需要执行

    dts-gen -m <module_name>

然后在 tsconfig.json 文件里加上一个属性

    "include": [
        ...,
        ...,
        "<module_name>.d.ts",
    ]

很恼人啊！