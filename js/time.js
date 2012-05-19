var elapsedTime = function(createdAt) {
    var ageInSeconds = (new Date().getTime() - new Date(createdAt).getTime()) / 1000;
    if (ageInSeconds < 2) {
        return '刚刚';
    }
    if (ageInSeconds < 60) {
        var n = ageInSeconds;
        return Math.floor(n) + '秒前';
    }
    if (ageInSeconds < 60 * 60) {
        var n = Math.floor(ageInSeconds/60);
        return n + '分钟前';
    }
    if (ageInSeconds < 60 * 60 * 24) {
        var n = Math.floor(ageInSeconds/60/60);
        return n + '小时前';
    }
    if (ageInSeconds < 60 * 60 * 24 * 7) {
        var n = Math.floor(ageInSeconds/60/60/24);
        return n + '日前';
    }
    if (ageInSeconds < 60 * 60 * 24 * 31) {
        var n = Math.floor(ageInSeconds/60/60/24/7);
        return n + '星期前';
    }
    if (ageInSeconds < 60 * 60 * 24 * 365) {
        var n = Math.floor(ageInSeconds/60/60/24/31);
        return n + '个月前';
    }
    var n = Math.floor(ageInSeconds/60/60/24/365);
    return n + '年前';
}

function fixDate(d) {
    var a = d.split(' ');
    var year = a.pop();
    return a.slice(0, 3).concat([year]).concat(a.slice(3)).join(' ');
}
