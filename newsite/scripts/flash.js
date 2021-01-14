function flash(op) {

    var fvars = (op.vars) ? op.vars : '';
    var color = (op.color) ? op.color : '#ffffff';

    html = '<object classid="clsid:d27cdb6e-ae6d-11cf-96b8-444553540000" codebase="http://fpdownload.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=7,0,0,0" width="' + op.w + '" height="' + op.h + '" align="middle">';
    html += '<param name="allowScriptAccess" value="sameDomain" />';
    html += '<param name="movie" value="' + op.src + '" />';
    html += '<param name="quality" value="high" />';
    html += '<param name="bgcolor" value="' + color + '" />';
    html += '<param name="flashvars" value="' + fvars + '" />';
    html += '<video width="' + op.w + '" height="' + op.h + '" autoplay>';
    html += '<source src="' + op.src + '" type="video/mp4">';
    // html += '<source src="' + op.src + '" type="video/ogg">';
    html += '</video>';
    // html += '<embed src="' + op.src + '" flashvars="' + fvars + '" quality="high" bgcolor="' + color + '" width="' + op.w + '" height="' + op.h + '" name="top" align="middle" allowScriptAccess="sameDomain" type="application/x-shockwave-flash" pluginspage="http://www.macromedia.com/go/getflashplayer" />';
    html += '</object>';

    document.write(html);
}
