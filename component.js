<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <meta name="description" content="接单与出货双模块、人民币与原币双口径差异核对">
  <title>接单与出货差异核对</title>
  <link rel="stylesheet" href="/styles.css">
</head>
<body>
  <header class="topbar">
    <div class="brand">
      <span class="brand-mark">P</span>
      <div><strong>POOWARD</strong><small>市场部数据核对中心</small></div>
    </div>
    <div class="topbar-status">
      <span class="module-badge" id="moduleBadge">接单差异核对</span>
      <span class="basis" id="basisBadge">人民币口径 · RMB</span>
    </div>
  </header>

  <div class="workspace">
    <aside class="sidebar" aria-label="核对模块">
      <div class="sidebar-title">核对模块</div>
      <button class="nav-item active" type="button" data-module="order">
        <span class="nav-number">01</span>
        <span><strong>接单差异</strong><small>Order reconciliation</small></span>
      </button>
      <button class="nav-item" type="button" data-module="shipment">
        <span class="nav-number">02</span>
        <span><strong>出货差异</strong><small>Shipment reconciliation</small></span>
      </button>
      <div class="sidebar-note">
        <strong>每次只核对一个模块</strong>
        <span>上传文控表和当前模块的两份系统表，人民币与原币结果会同时计算并缓存。</span>
      </div>
    </aside>

    <main class="shell">
      <section class="hero">
        <div>
          <p class="eyebrow" id="moduleEyebrow">ORDER RECONCILIATION</p>
          <h1 id="moduleTitle">接单差异核对</h1>
          <p id="basisRule">系统金额 = 接单金额(RMB) + 出货运费(RMB)；文控金额 = “接单”子表的 VAT PRICE。</p>
        </div>
        <div class="formula"><span>系统</span><b>−</b><span>文控</span><b>=</b><em>差异</em></div>
      </section>

      <section class="panel upload-panel" id="uploadPanel">
        <div class="section-title">
          <div><span class="step">01</span><h2 id="uploadTitle">上传接单核对三份文件</h2></div>
          <p>支持点击选择或拖拽上传 .xlsx / .xlsm；文件仅在运行工具的电脑内存中处理</p>
        </div>
        <form id="uploadForm">
          <section class="shared-source">
            <div class="group-heading">
              <span class="group-kicker">SHARED SOURCE</span>
              <div><h3>共用文控数据</h3><p>切换接单或出货模块时，已选择的文控文件会继续保留。</p></div>
            </div>
            <label class="drop-card shared-card">
              <span class="file-icon">文</span>
              <span class="file-title">文控登记表 <small>Document Control Ledger</small></span>
              <span class="file-rule">接单：VAT PRICE / TP-CPO<br>出货：VAT PRICE / TP-CPO</span>
              <input required type="file" name="document" accept=".xlsx,.xlsm">
              <span class="pick">选择或拖入</span><span class="filename">尚未选择</span>
            </label>
          </section>

          <div class="source-groups">
            <section class="source-group active" data-source-module="order">
              <div class="group-heading">
                <span class="group-kicker">ORDER</span>
                <div><h3>接单核对数据</h3><p>用于核对文控表中名称含“接单”的子表。</p></div>
              </div>
              <div class="upload-grid two-cols">
                <label class="drop-card">
                  <span class="file-icon">接</span>
                  <span class="file-title">接单金额明细 <small>Order Amount Detail</small></span>
                  <span class="file-rule" id="orderAmountRule">读取接单金额(RMB)</span>
                  <input type="file" name="order_amount" accept=".xlsx,.xlsm">
                  <span class="pick">选择或拖入</span><span class="filename">尚未选择</span>
                </label>
                <label class="drop-card">
                  <span class="file-icon">运</span>
                  <span class="file-title">接单运费明细 <small>Order Freight Detail</small></span>
                  <span class="file-rule" id="orderFreightRule">追加出货运费(RMB)</span>
                  <input type="file" name="order_freight" accept=".xlsx,.xlsm">
                  <span class="pick">选择或拖入</span><span class="filename">尚未选择</span>
                </label>
              </div>
            </section>

            <section class="source-group" data-source-module="shipment">
              <div class="group-heading">
                <span class="group-kicker">SHIPMENT</span>
                <div><h3>出货核对数据</h3><p>用于核对文控表中名称含“出货”的子表。</p></div>
              </div>
              <div class="upload-grid two-cols">
                <label class="drop-card">
                  <span class="file-icon">货</span>
                  <span class="file-title">出货金额明细 <small>Shipment Amount Detail</small></span>
                  <span class="file-rule" id="shipmentAmountRule">读取出货金额(RMB)</span>
                  <input type="file" name="shipment_amount" accept=".xlsx,.xlsm">
                  <span class="pick">选择或拖入</span><span class="filename">尚未选择</span>
                </label>
                <label class="drop-card">
                  <span class="file-icon">费</span>
                  <span class="file-title">出货运费明细 <small>Shipment Freight Detail</small></span>
                  <span class="file-rule" id="shipmentFreightRule">追加出货运费(RMB)</span>
                  <input type="file" name="shipment_freight" accept=".xlsx,.xlsm">
                  <span class="pick">选择或拖入</span><span class="filename">尚未选择</span>
                </label>
              </div>
            </section>
          </div>

          <div class="action-row">
            <fieldset class="basis-control">
              <legend>结果显示口径</legend>
              <div class="segmented">
                <label><input type="radio" name="basis" value="rmb" checked><span>人民币 RMB</span></label>
                <label><input type="radio" name="basis" value="original"><span>原币</span></label>
              </div>
              <small class="cache-note">提交后切换口径无需重新计算</small>
            </fieldset>
            <div class="right-actions">
              <label class="tolerance"><span id="toleranceLabel">差异容差（元）</span><input type="number" name="tolerance" value="1.00" min="0" step="0.01"></label>
              <button class="primary" id="analyzeButton" type="submit"><span>开始接单核对</span><span class="arrow">→</span></button>
            </div>
          </div>
        </form>
        <div class="progress hidden" id="progress"><span></span><p id="progressText">正在读取接单模块三份文件，并生成人民币与原币结果…</p></div>
        <div class="error hidden" id="errorBox"></div>
      </section>

      <section id="results" class="hidden results">
        <div class="result-head">
          <div><p class="eyebrow" id="resultEyebrow">ORDER · RMB</p><h2 id="resultTitle">接单差异核对结果</h2></div>
          <button id="downloadButton" class="secondary">下载当前模块 Excel 结果</button>
        </div>
        <div class="cache-status"><span class="cache-dot"></span><span id="cacheStatusText">当前模块的人民币与原币结果均已缓存，可直接切换口径。</span></div>
        <div class="kpi-grid" id="kpis"></div>
        <div id="warningBox" class="warning-box hidden"></div>

        <section class="panel table-panel">
          <div class="section-title compact">
            <div><span class="step">02</span><h2>客户差异汇总</h2></div>
            <label class="search">搜索客户代码<input id="customerSearch" type="search" placeholder="如 CB229"></label>
          </div>
          <p class="hint">点击客户代码，查看该客户有差异的订单流水号。</p>
          <div class="table-wrap"><table id="customerTable"></table></div>
        </section>

        <section class="panel table-panel hidden" id="linePanel">
          <div class="section-title compact">
            <div><span class="step">03</span><h2><span id="lineCustomer"></span> 的差异流水号</h2></div>
            <span class="count" id="lineCount"></span>
          </div>
          <p class="hint">点击订单流水号，同时查看文控登记表与当前系统模块的原始行。</p>
          <div class="table-wrap"><table id="lineTable"></table></div>
        </section>

        <section class="panel detail-panel hidden" id="detailPanel">
          <div class="section-title compact">
            <div><span class="step">04</span><h2>流水号 <span id="detailOrder"></span> 原始数据</h2></div>
            <span class="count" id="detailStatus"></span>
          </div>
          <div class="comparison">
            <article>
              <div class="source-head"><span class="source-dot doc"></span><h3>文控登记表</h3><span id="docCount"></span></div>
              <div id="docDetails" class="raw-list"></div>
            </article>
            <article>
              <div class="source-head"><span class="source-dot sys"></span><h3 id="systemDetailTitle">系统接单数据</h3><span id="sysCount"></span></div>
              <div id="sysDetails" class="raw-list"></div>
            </article>
          </div>
        </section>
      </section>
    </main>
  </div>

  <footer>POOWARD · order & shipment reconciliation utility</footer>
  <script src="/app.js"></script>
</body>
</html>
