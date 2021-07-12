function form_val(i) {
    const items = i_vue.content[cur_key];
    i = cur_key;
    $('.count').html(`当前数量：${i + 1}`);
    $('form .company').val(items['企业名称'] || '');
    $('form .customer').val(items['法定代表人'] || '');
    $('form .phonenum').val(items['联系电话'] || '');
    $('form .area').val(items['省份'] + items['城市'] || '');
    $('form .address').val(items['地址'] || '');
    $('form .email').val(items['邮箱'] || '');
    $('form .company_website').val(items['网址'] || '');
    $('form .industry').val(items['经营范围'].split(';')[0] || '');
    $('form .established').val(items['成立日期'] || '');
    $('form .num_of_p').val(items['参保人数'] || '');
    $('form .comments').val('');
    $('form .salesperson').val(items['销售人员'] || '');

    let cmets = [];
    if (items['留言']) {
        cmets = items['留言'].split("\n").reverse();
        // console.log(cmets);
        if (cmets.length > 8) {
            cmets.splice(8, cmets.length);
        }
    }
    $('form .history_comments').val(
        cmets.join('\n').trim()
    );
}

function save_form(callback) {
    return new Promise((resolve, reject) => {
        const e_sales_v = e_salesperson.val();
        const e_cmets_v = e_comments.val();
        if (!i_vue.content) { alert("先导入Excel"); return; }

        if (e_sales_v && e_cmets_v) {
            const time = new Date();
            const timestr = `${time.getFullYear()}/${time.getMonth() + 1}/${time.getDate()}`;

            const str = `${timestr} : 售后人员: ${e_salesperson.val()}, 留言内容: ${e_comments.val()};`

            let comments = i_vue.content[cur_key]["留言"];
            i_vue.content[cur_key]["留言"] = !comments ? `${str}` : `${comments} \n ${str}`;
            e_comments.val('');
            e_salesperson.val('');
        }

        // console.log(i_vue.content);
        if (callback) {
            callback(reject);
        }

        resolve(true);
    }).catch((e) => {
        // console.log(e);
    })
}

var cur_key = 0;
var e_comments, e_salesperson;

$(() => {
    e_comments = $(".comments");
    e_salesperson = $(".salesperson");

    document.getElementById('file').addEventListener('change', async function (e) {
        setTimeout(() => {
            cur_key = 0;
            form_val();
        }, 1000);
    });

    $('.pre_one').on('click', async () => {
        const succ = await save_form((reject) => {
            if (cur_key <= 0) {
                reject(false);
            }
        })

        if (succ) {
            cur_key--;
            form_val();
        } else {
            alert("已到最前一个");
        }
    })

    $('.next_one').on('click', async () => {
        const succ = await save_form((reject) => {
            if (cur_key >= i_vue.content.length - 1) {
                reject(false);
            }
        })

        if (succ) {
            cur_key++;
            form_val();
        } else {
            alert("已到最后一个");
        }
    })
})
