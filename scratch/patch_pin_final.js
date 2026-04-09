const fs = require('fs');
const content = fs.readFileSync('docs/order.html', 'utf8');

// The snippet we want to replace
const oldSnippet = `    if (data.activeOrders) {
       S.activeOrders = data.activeOrders;
    }
    
    // Transition to PIN entry
    sLoading(false); 
    
    if (data.hasPin) {
       S.authMode = \"verify\";
       const parts = (data.name || \"\").trim().split(\" \");
       const firstName = parts[0] || \"\";
       $(\"pinAuthTitle\").innerHTML = firstName ? \`🔒 Welcome back, \${firstName}\` : \"🔒 Welcome back\";
       $(\"pinAuthDesc\").innerHTML = \"Please enter your 4-digit PIN to access your account.\";
       if($(\"pinHint\")) $(\"pinHint\").style.display = \"none\";
       $(\"forgotPinWrap\").style.display = \"block\";
       $(\"pinAuthSection\").style.display = \"block\";
       $(\"btnNext\").textContent = \"Verify PIN & Continue\";
       $(\"btnNext\").disabled = false;
       setTimeout(()=>$(\"pinInput\").focus(), 100);
       return;
    } else {
       S.authMode = \"create\";
       $(\"pinAuthTitle\").innerHTML = \"🔒 Protect your account\";
       $(\"pinAuthDesc\").innerHTML = \"Please set a new 4-digit secure PIN for your Wallet.\";
       if($(\"pinHint\")) $(\"pinHint\").style.display = \"block\";
       $(\"forgotPinWrap\").style.display = \"none\";
       $(\"pinAuthSection\").style.display = \"block\";
       $(\"btnNext\").textContent = \"Set PIN & Continue\";
       $(\"btnNext\").disabled = false;
       S.tempProfileData = data.found ? data : null;
       setTimeout(()=>$(\"pinInput\").focus(), 100);
       return;
    }`;

const newSnippet = `    if (data.activeOrders) {
       S.activeOrders = data.activeOrders;
    }
    
    // Transition to PIN entry
    sLoading(false); 
    
    if (data.hasPin) {
       S.authMode = \"verify\";
       const parts = (data.name || \"\").trim().split(\" \");
       const firstName = parts[0] || \"\";
       $(\"pinAuthTitle\").innerHTML = firstName ? \`🔒 Welcome back, \${firstName}\` : \"🔒 Welcome back\";
       $(\"pinAuthDesc\").innerHTML = \"Please enter your 4-digit PIN to access your account.\";
       if($(\"pinHint\")) $(\"pinHint\").style.display = \"none\";
       $(\"forgotPinWrap\").style.display = \"block\";
       $(\"pinAuthSection\").style.display = \"block\";
       $(\"btnNext\").textContent = \"Verify PIN & Continue\";
       $(\"btnNext\").disabled = false;
       setTimeout(()=>$(\"pinInput\").focus(), 100);
       return;
    } else {
       S.authMode = \"create\";
       if (data.found) {
          const parts = (data.name || \"\").trim().split(\" \");
          const firstName = parts[0] || \"\";
          $(\"pinAuthTitle\").innerHTML = firstName ? \`🔒 Welcome back, \${firstName}\` : \"🔒 Welcome back\";
          $(\"pinAuthDesc\").innerHTML = \"Please set a <strong>new</strong> 4-digit secure PIN.\";
          S.tempProfileData = data;
       } else {
          $(\"pinAuthTitle\").innerHTML = \"🔒 Protect your account\";
          $(\"pinAuthDesc\").innerHTML = \"Please set a 4-digit secure PIN for your Wallet.\";
          S.tempProfileData = { _newAccount: true };
       }
       if($(\"pinHint\")) $(\"pinHint\").style.display = \"block\";
       $(\"forgotPinWrap\").style.display = \"none\";
       $(\"pinAuthSection\").style.display = \"block\";
       $(\"btnNext\").textContent = \"Set PIN & Continue\";
       $(\"btnNext\").disabled = false;
       setTimeout(()=>$(\"pinInput\").focus(), 100);
       return;
    }`;

// We need to be careful with exact whitespace. Let's use a simpler match if possible or normalize.
// Since we have the file content in 'content', let's use a more robust replacement.

// Note: The previous view showed lines around 1640+ for this logic.
// Let's try to find a unique part of the snippet.

const searchStr = 'S.authMode = \"create\";';
if (content.indexOf(searchStr) === -1) {
    console.error(\"COULD NOT FIND searchStr\");
    process.exit(1);
}

// Replacement logic: replace the entire 'else { S.authMode = \"create\" ... }' block.
const elseBlockRegex = /else \{\s+S\.authMode = \"create\";\s+\$\(\"pinAuthTitle\"\)\.innerHTML = \"🔒 Protect your account\";\s+\$\(\"pinAuthDesc\"\)\.innerHTML = \"Please set a new 4-digit secure PIN for your Wallet\.\";\s+if\(\$\(\"pinHint\"\)\) \$\(\"pinHint\"\)\.style\.display = \"block\";\s+\$\(\"forgotPinWrap\"\)\.style\.display = \"none\";\s+\$\(\"pinAuthSection\"\)\.style\.display = \"block\";\s+\$\(\"btnNext\"\)\.textContent = \"Set PIN & Continue\";\s+\$\(\"btnNext\"\)\.disabled = false;\s+S\.tempProfileData = data\.found \? data : null;\s+setTimeout\(\(\)=> \$\(\"pinInput\"\)\.focus\(\), 100\);\s+return;\s+\}/;

if (!elseBlockRegex.test(content)) {
    console.error(\"COULD NOT MATCH elseBlockRegex\");
    // Backup: just replace the specific lines
    const content2 = content.replace('S.tempProfileData = data.found ? data : null;', 'S.tempProfileData = data.found ? data : { _newAccount: true };');
    fs.writeFileSync('docs/order.html', content2);
    console.log(\"FALLBACK REPLACEMENT DONE\");
} else {
    const finalContent = content.replace(elseBlockRegex, \`else {
       S.authMode = \"create\";
       if (data.found) {
          const parts = (data.name || \"\").trim().split(\" \");
          const firstName = parts[0] || \"\";
          $(\"pinAuthTitle\").innerHTML = firstName ? \\\`🔒 Welcome back, \\\${firstName}\\\` : \"🔒 Welcome back\";
          $(\"pinAuthDesc\").innerHTML = \"Please set a <strong>new</strong> 4-digit secure PIN.\";
          S.tempProfileData = data;
       } else {
          $(\"pinAuthTitle\").innerHTML = \"🔒 Protect your account\";
          $(\"pinAuthDesc\").innerHTML = \"Please set a 4-digit secure PIN for your Wallet.\";
          S.tempProfileData = { _newAccount: true };
       }
       if($(\"pinHint\")) $(\"pinHint\").style.display = \"block\";
       $(\"forgotPinWrap\").style.display = \"none\";
       $(\"pinAuthSection\").style.display = \"block\";
       $(\"btnNext\").textContent = \"Set PIN & Continue\";
       $(\"btnNext\").disabled = false;
       setTimeout(()=>$(\"pinInput\").focus(), 100);
       return;
    }\`);
    fs.writeFileSync('docs/order.html', finalContent);
    console.log(\"FULL REPLACEMENT DONE\");
}
