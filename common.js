"use strict";

const arr_avg = numbers => {
    let mean = 0
    if (numbers.length > 0) {
        mean = numbers.reduce((acc, n) => acc + n, 0) / numbers.length; 
    }
    return mean;
}

const standardDeviation_avg = numbers => {
    numbers = numbers.filter(n => n)        // выбраковка пустых элементов
    let mean = arr_avg(numbers)
    let std_dev = 0
    if (numbers.length>1) {
        std_dev = Math.sqrt(
                numbers.reduce((acc, n) => (n - mean) ** 2) / (numbers.length - 1)
            )
        }
    return [
        mean, std_dev
        ];
};

function apply_offset(coords, offset) {
    return { c: coords.c + offset.c, r: coords.r + offset.r };
}


const ZERO_OFFSET = { c: 0, r: 0 };

function getVarRangeFromHeader(sheet, str_range, id_offset, name_offset) {
    let range = XLSX.utils.decode_range(str_range);

    let variable_dimension = range.s.c === range.e.c ? "r" : "c";
    let res = { variable_dimension };
    res.header_range = range;
    res.def = [];

    let i = {};
    let counter = 0;
    let prev_start = range.s[variable_dimension];
    for (i.c = range.s.c; i.c <= range.e.c; ++i.c) {
        for (i.r = range.s.r; i.r <= range.e.r; ++i.r) {
            let data = sheet[XLSX.utils.encode_cell(i)];

            if (data && data.v) {
                counter++;
                let new_def = { 's': i[variable_dimension] };
                new_def.id = id_offset
                    ? sheet[XLSX.utils.encode_cell(apply_offset(i, id_offset))].v
                    : counter;
                if (name_offset) {
                    new_def.name =
                        sheet[XLSX.utils.encode_cell(apply_offset(i, name_offset))].v;
                    new_def.name = new_def.name.replace(/\s+/g,' ').trim();
                }
                res.def.push(new_def);
                prev_start = i[variable_dimension];
            }
        }
    }
    let ids = {}
    for (let i=0; i<res.def.length;i++) {
        ids[res.def[i].id] = i;
    }
    res.ids = ids;
    return res;
}

function readWSData(sheet, criteria, subjects, keys_order) {
    if (criteria.variable_dimension !== subjects.variable_dimension) {

        let reversed_addr = criteria.variable_dimension==='r';

        let dict = {}       // словарь значений внутри пересечения критерия и дисцпилины
        let data = []       // данные на выходе

        let i = {}
        // let current_criterion_idx = 0, current_subject_idx = 0;
        for (let c_i = criteria.header_range.s[criteria.variable_dimension], current_criterion_idx = -1; c_i<= criteria.header_range.e[criteria.variable_dimension]; c_i++) {
            if (current_criterion_idx<criteria.def.length-1 && c_i === criteria.def[current_criterion_idx+1].s) {
                // console.log("CR", criteria.def[current_criterion_idx])

                current_criterion_idx+=1;
            }
            for (let s_i = subjects.header_range.s[subjects.variable_dimension], current_subject_idx = -1; s_i<= subjects.header_range.e[subjects.variable_dimension]; s_i++) {
                if (current_subject_idx<subjects.def.length-1 && s_i === subjects.def[current_subject_idx+1].s) {
                    current_subject_idx+=1;
                }
                i[criteria.variable_dimension] = c_i; 
                i[subjects.variable_dimension] = s_i; 
                //console.log("I",i)
                let value = 0
                try {
                    value = parseInt(sheet[XLSX.utils.encode_cell(i)].v)
                    if (value === 0) {
                        value = 1;
                    }
                } catch {
                    /*...*/
                } 
                if (value) {
                    console.log('V',value, 'C', current_criterion_idx, 'S', current_subject_idx)
                    let key = String([current_criterion_idx, current_subject_idx])
                    if (key in dict) {
                        dict[key].push(value)
                        if (dict[key].length == keys_order.length) {
                            let new_item = {'crit': current_criterion_idx, 'subj': current_subject_idx }
                            for (let j=0; j<dict[key].length;j++) {
                                new_item[keys_order[j]] = dict[key][j];
                            }
                        data.push(new_item)
                        }
                    } else {
                        dict[key] = [value]
                    }
                }
            }
        }
        return data;
    }
}

function SaveSetting(key, value) {
    localStorage.setItem(key, value)
}


function copyTableToClipboard() {
    if (tblOut.innerHTML) {
        let sel = window.getSelection()
        sel.removeAllRanges()
        let range = new Range()
        let r_container = tblOut.firstElementChild;
        console.log(r_container)
        range.setStartBefore(r_container.firstElementChild)
        range.setEndAfter(r_container.lastElementChild.lastElementChild)
        sel.addRange(range)
        document.execCommand('copy')
        sel.removeAllRanges()
    }
}

function exportToWord(tag) {

   let html, link, blob, url, css;

   css = (
     '<style>' +
     '@page WordSection1{size: 595.35pt 841.95pt;mso-page-orientation: portrait;}' +
     'div.WordSection1 {page: WordSection1;}' +
     '</style>'
   );

   html = tag.innerHTML;
   blob = new Blob(['\ufeff', css + html], {
     type: 'application/msword'
   });
   url = URL.createObjectURL(blob);
   link = document.createElement('A');
   link.href = url;
   link.download = 'Вывод.doc'; 
   link.click();  
 };

 function triggerExport() {
    exportToWord(output)
 }