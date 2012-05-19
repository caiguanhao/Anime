var elapsedTime = function(createdAt) {
    var ageInSeconds = (new Date().getTime() - new Date(createdAt).getTime()) / 1000;
    if (ageInSeconds < 2) {
        return '�ո�';
    }
    if (ageInSeconds < 60) {
        var n = ageInSeconds;
        return Math.floor(n) + '��ǰ';
    }
    if (ageInSeconds < 60 * 60) {
        var n = Math.floor(ageInSeconds/60);
        return n + '����ǰ';
    }
    if (ageInSeconds < 60 * 60 * 24) {
        var n = Math.floor(ageInSeconds/60/60);
        return n + 'Сʱǰ';
    }
    if (ageInSeconds < 60 * 60 * 24 * 7) {
        var n = Math.floor(ageInSeconds/60/60/24);
        return n + '��ǰ';
    }
    if (ageInSeconds < 60 * 60 * 24 * 31) {
        var n = Math.floor(ageInSeconds/60/60/24/7);
        return n + '����ǰ';
    }
    if (ageInSeconds < 60 * 60 * 24 * 365) {
        var n = Math.floor(ageInSeconds/60/60/24/31);
        return n + '����ǰ';
    }
    var n = Math.floor(ageInSeconds/60/60/24/365);
    return n + '��ǰ';
}

function fixDate(d) {
    var a = d.split(' ');
    var year = a.pop();
    return a.slice(0, 3).concat([year]).concat(a.slice(3)).join(' ');
}
