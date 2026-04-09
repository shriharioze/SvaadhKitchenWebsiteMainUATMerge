const fs = require('fs'); 
let c = fs.readFileSync('docs/order.html', 'utf8'); 

const oldTitle = '        $("pinAuthTitle").innerHTML = "🔒 Protect your account";';
const newTitle = '        if (data.found) {\n           const parts = (data.name || "").trim().split(" ");\n           const firstName = parts[0] || "";\n           $("pinAuthTitle").innerHTML = firstName ? "🔒 Welcome back, " + firstName : "🔒 Welcome back";\n        } else {\n           $("pinAuthTitle").innerHTML = "🔒 Protect your account";\n        }';

const oldDesc = '        $("pinAuthDesc").innerHTML = "Please set a new 4-digit secure PIN for your Wallet.";';
const newDesc = '        if (data.found) {\n           $("pinAuthDesc").innerHTML = "Please set a <strong>new</strong> 4-digit secure PIN.";\n        } else {\n           $("pinAuthDesc").innerHTML = "Please set a 4-digit secure PIN for your Wallet.";\n        }';

const oldData = '        S.tempProfileData = data.found ? data : null;';
const newData = '        S.tempProfileData = data.found ? data : { _newAccount: true };';

c = c.replace(oldTitle, newTitle);
c = c.replace(oldDesc, newDesc);
c = c.replace(oldData, newData);

fs.writeFileSync('docs/order.html', c);
console.log("REPLACEMENT SUCCESSFUL");
