const fs = require('fs');
const content = fs.readFileSync('docs/order.html', 'utf8');

const targetStr = 'S.tempProfileData = data.found ? data : null;';
const replacementStr = 'if (data.found) {\n          const parts = (data.name || "").trim().split(" ");\n          const firstName = parts[0] || "";\n          $("pinAuthTitle").innerHTML = firstName ? `🔒 Welcome back, ${firstName}` : "🔒 Welcome back";\n          $("pinAuthDesc").innerHTML = "Please set a <strong>new</strong> 4-digit secure PIN.";\n          S.tempProfileData = data;\n       } else {\n          $("pinAuthTitle").innerHTML = "🔒 Protect your account";\n          $("pinAuthDesc").innerHTML = "Please set a 4-digit secure PIN for your Wallet.";\n          S.tempProfileData = { _newAccount: true };\n       }';

if (content.indexOf(targetStr) === -1) {
    console.log("NOT FOUND");
} else {
    // Only replace the first occurrence (which is inside the 'create' block)
    const result = content.replace(targetStr, replacementStr);
    fs.writeFileSync('docs/order.html', result);
    console.log("SUCCESS");
}
