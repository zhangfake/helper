// ==UserScript==
// @name        exportor
// @namespace   Violentmonkey Scripts
// @match       *://crm-taobao.huanleguang.com/*
// @grant       GM_xmlhttpRequest
// @grant       GM_setValue
// @grant       GM_getValue
// @version     1.0
// @author      -
// @description 2023/6/19 18:18:00
// @run-at      document-idle
// @require     https://unpkg.com/layui@2.8.7/dist/layui.js
// @require     https://unpkg.com/xlsx/dist/xlsx.full.min.js
// ==/UserScript==

const styleSheet = document.createElement('link');
styleSheet.setAttribute("href", "https://unpkg.com/layui@2.8.7/dist/css/layui.css")
styleSheet.setAttribute("rel", "stylesheet")
document.body.appendChild(styleSheet)

const innerStyle = document.createElement("style");
innerStyle.innerHTML = `
.plugin-menu-button {
  position: fixed;
  left: 0;
  top: 0;
  cursor: pointer;
  z-index: 99999
}
textarea {
  resize: none
}
`
document.body.appendChild(innerStyle)


const login = (user, password) => {
  return new Promise((resolve, reject) => {
    GM_xmlhttpRequest({
      url: "https://crm-taobao.huanleguang.com/user/login",
      method: "post",
      responseType: "json",
      headers: {
        "Content-Type": "application/json"
      },
      data: JSON.stringify({user, password}),
      onload: resolve,
      onerror: reject
    })
  })
}

const getUserCookie = (url) => {
  return new Promise((resolve, reject) => {
    GM_xmlhttpRequest({
      url: url,
      method: "get",
      onload: resolve,
      onerror: reject
    })
  })
}

const getRows = (page, pageSize, config, cookie) => {
  // config.start_time = '2023-05-30 00:00:00'
  // config.end_time = '2023-05-30 23:59:59'
  return new Promise((resolve, reject) => {
    GM_xmlhttpRequest({
      url: "https://crm-taobao.huanleguang.com/son/get-data?start_time="+config.start_time+"&end_time="+config.end_time+"&page_size="+pageSize+"&page="+page,
      headers: {
        cookie
      },
      responseType: "json",
      method: "get",
      onload: resolve,
      onerror: reject
    })
  })
}

const getSonParents = (cookie) => {
return new Promise((resolve, reject) => {
    GM_xmlhttpRequest({
      url: "https://crm-taobao.huanleguang.com/son/parent",
      headers: {
        cookie
      },
      responseType: "json",
      method: "get",
      onload: resolve,
      onerror: reject
    })
  })
}


const gotoDataReport = () => {
  location.hash = '#/data-report'
}
const gotoLogin = () => {
  location.hash = '#/login'
}

const formatDateTime = (timestamp) => {
  const date = new Date(timestamp);
  return [date.getFullYear(), (date.getMonth()+1).toString().padStart(2,'0'), date.getDate().toString().padStart(2,'0')].join('-')
    + ' ' + [date.getHours().toString().padStart(2,'0'), date.getMinutes().toString().padStart(2,'0'), date.getSeconds().toString().padStart(2,'0')].join(':')
}


layui.use(function(){
  const layer = layui.layer;
  const laydate = layui.laydate;
  const form = layui.form;

  let config = JSON.parse(GM_getValue('config', '{}'))
  const startDownloadThread = async field => {
    // 存储当前数据
    config = field
    field.end_time = formatDateTime(Date.now())
    GM_setValue("config", JSON.stringify(field))

    // 处理账号列表
    const accounts = field.users.split(/\n/g).map(item => {
      let [username, password] = item.split('-')
      return {username, password}
    }).filter(item => item.username)

    let datas = []
    for (let i = 0; i < accounts.length; i++ ){
      const account = accounts[i]
      gotoLogin()
      const result = await login(account.username, account.password);

      layer.msg('正在登录账号:'+account.password, {icon: 1});
      const input = document.querySelectorAll('.hlg-input')
      input[0].value = account.username
      input[1].value = account.password

      if (result.response.code == 10000) {
          const res = await getUserCookie(result.response.data)
          layer.msg('登录成功:' + account.password, {icon: 1});
          cookie = res.responseHeaders.match(/set-cookie: ([^;]*)/)[1]
          gotoDataReport()
          const sonParents = await getSonParents()
          const providerMap = Object.fromEntries(sonParents.response.data.map(item => ([item.relate_id, item.name])))
          let page = 1,totalPage = 1, pageSize=1000
          do {
            const list = await getRows(page++, pageSize, field, cookie)
            datas = datas.concat(list.response.data.map(item => {
              item.supplier_name = providerMap[item.supplier] || ''
              item.trade_info = item.trade ? JSON.parse(item.trade) : {}
              return item
            }))
            totalPage = Math.ceil(list.response.total / pageSize)
          } while(page < totalPage);
      } else {
       layer.alert(result.response.message, {
          title: '账号'+account.password+'登录失败'
        });
      }
    }

    if (!datas.length) {
      layer.msg('没有数据可以导出', {icon: 1});
      return;
    }
    let csv = datas.map(item => {
      return {"标识":item?.params?.bs || '',"来源": item.supplier_name || '',
              "商品名称": item.item_title || '', "商品ID": item.item_id || '',"用户行为": {
  cart: "加购",
  collect: "收藏",
  coupon: "领券",
  toBuy: "跳转购买",
  tradeClose: "订单关闭",
  tradeCreate: "创建订单",
  tradePatlyPay: "预售单付款",
  tradePatlyRefund: "部分退款",
  tradePay: "订单付款",
  tradeRefundSuccess: "退款成功",
  tradeSellerShip: "已发货",
  tradeSuccess: "订单成功",
  view: "浏览"
}[item.event] || '', "操作时间": formatDateTime(item.event_time*1000),
              "订单状态": {
                  WAIT_BUYER_PAY: '等待买家付款',
                  WAIT_SELLER_SEND_GOODS: '等待发货',
                  WAIT_BUYER_CONFIRM_GOODS: '已发货',
                  TRADE_FINISHED: '需要评价',
                  TRADE_PARTLY_REFUND: '部分退款',
                  TRADE_CLOSED: '已退款',
              }[item.trade_info.orders_status] || '',"订单ID": item.trade_info.tid || '',"预售/订单金额": item.trade_info.payment || '',"订单时间":item.trade_info.created_time || '',
               "详细数据":item.params ? Object.keys(item.params).map((t=>`${t}: ${item.params[t]}`)).join("\r\n") : ""}
    })

    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(csv, { header: ["标识","来源","商品名称","商品ID","用户行为","操作时间","订单状态","订单ID","预售/订单金额","订单时间","详细数据"] });
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sheet1");
    // 将工作簿保存为 Excel 文件
    XLSX.writeFile(workbook, "data.xlsx");
  }

  const button = document.createElement('div')
  button.setAttribute('class', 'plugin-menu-button')
  button.innerHTML = '<i class="layui-icon layui-icon-menu-fill" style="font-size: 30px; color: #1E9FFF;"></i>'

  button.onclick = () => {
      layer.open({
        title: "批量下载",
        type: 1,
        offset: 'l',
        anim: 'slideRight', // 从左往右
        area: ['320px', '100%'],
        shade: 0.1,
        shadeClose: true,
        id: 'ID-demo-layer-direction-l',
        success: function(){
          laydate.render({
            elem: '#ID-laydate-type-datetime',
            type: 'datetime',
            fullPanel: true // 2.8+
          });

          form.on('submit(demo1)', function(data){
            var field = data.field; // 获取表单字段值

            startDownloadThread(field);
            return false
          });

        },
        content: `
<div style="padding: 16px;">
  <form class="layui-form" action="">
    <div class="layui-form-item">
      <div>账号列表</div>
      <textarea style="resize: none;" name="users" placeholder="每行一个账号使用-分割账号密码 例: 18688888888-123456" rows="10" class="layui-textarea">${config.users}</textarea>
    </div>
     <div class="layui-form-item">
      <div>开始时间</div>
       <input type="text" class="layui-input" value="${config.end_time}" name="start_time" id="ID-laydate-type-datetime" placeholder="yyyy-MM-dd HH:mm:ss">
    </div>
      <div class="layui-form-item">
    <div class="layui-input-block">
      <button type="submit" class="layui-btn" lay-submit lay-filter="demo1">开始下载</button>
    </div>
  </div>
</div>`
      });

  }

  document.body.appendChild(button)

});
