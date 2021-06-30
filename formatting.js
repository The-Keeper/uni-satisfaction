"use strict";
/* возвращает отформатированное число num с точностью до presc знаков после запятой */
let presc = 2;

function f_p(num) {
    if (num !== undefined)
        return num.toFixed(presc).toString().replace("." , ",")
    else
        return ""
    };

const FORM = {
    'm':    { disc: n => `${f_p(n)}`, disc_sigma: n => `${f_p(n)}`, agr: n => `${f_p(n)}` },
    'p10':  { disc: n => `${f_p(n*10)}%`, disc_sigma: n => `${f_p(n*10)}`, agr: n => `${f_p(n*10)}` },
    'p100': { disc: n => `${f_p(n*100)}%`, disc_sigma: n => `${f_p(n*100)}`, agr: n => `${f_p(n**100)}` },
    'dist': { disc: n => `${f_p(n*100)}%`, disc_sigma: n => `${f_p(n)}`, agr: n => `${f_p(n)}` },
}