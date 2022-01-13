const nodemailer = require('nodemailer')
const xlsx = require('node-xlsx');
const fs = require('fs');


//创建一个SMTP客户端配置对象,这里以 qq 邮箱为例子
const transporter = nodemailer.createTransport({
    // 默认支持的邮箱服务包括：”QQ”、”163”、”126”、”iCloud”、”Hotmail”、”Yahoo”等
    service: "QQ",
    auth: {
        // 发件人邮箱账号
        user: 'xxx@qq.com',
        //发件人邮箱的授权码 需要在自己的邮箱设置中生成,并不是邮件的登录密码
        pass: 'xxxx'
    }
})

// 读取需要发送的数据
const sendDataXlsx = xlsx.parse("./" + "sendData.xlsx");
const sendData = sendDataXlsx[0].data
sendData.shift()
const errorMail=[ [
    '用户名',
    '邮箱',
    '序列号'
],];
const successMail=[
    [
        '用户名',
        '邮箱',
        '序列号'
    ],
];

sendData.map((item,index)=>{
    // 配置收件人信息
    const receiver = {
        // 发件人 邮箱  '昵称<发件人邮箱>'
        from: `"v2好友"<7443902@qq.com>`,
        // 主题
        subject: '红包封面序列号',
        // 收件人 的邮箱 可以是其他邮箱 不一定是qq邮箱
        to: item[1],
        // 可以使用html标签
        html: `
        <h2>${item[0]},你好</h2>
        <h1>红包封面序列号为:${item[2]}</h1>
    `
    }

    // 发送邮件
    transporter.sendMail(receiver, (error, info) => {
        if (error) {
            errorMail.push(item)
            writeToXlsx('sendFailed',successMail)
            return console.log('发送失败:', error);
        }
        transporter.close()
        successMail.push(item)
        writeToXlsx('sendSucceed',successMail)

    })
})

const writeToXlsx = (xlsxName,xlsxData) =>{
    // 写入发送xlsx
    const data = [
        {
            name : 'sheet1',
            data:xlsxData
        },
    ]
    const buffer = xlsx.build(data);
    fs.writeFile(`./${xlsxName}.xlsx`, buffer, function (err)
        {
            if (err){
                throw err;
            }
        }
    );
}
