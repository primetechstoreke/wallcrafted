(function () {
  "use strict";

  var EMAIL = "wallcrafted24@gmail.com";
  // Paste your Web3Forms access key here (get it free at https://web3forms.com)
  var ACCESS_KEY = "1b837676-7676-4ada-8465-3c865a055b20";
    var KEY = "wallcrafted_orders"; 
  var STATUSES = ["New", "In progress", "Ready", "Done"]; 

  var FIELDS = [
    ["name", "Name"], ["phone", "Phone"], ["email", "Email"], ["orderDate", "Order date"],
    ["design", "Design / item"], ["personalization", "Personalization text"],
    ["size", "Size / dimensions"], ["quantity", "Quantity"],
    ["colorStyle", "Color / style options"], ["baseType", "Base type"],
    ["instructions", "Special instructions"], ["price", "Price"],
    ["deposit", "Deposit paid"], ["paymentMethod", "Payment method"],
    ["completionDate", "Est. completion date"], ["pickup", "Pickup / shipping"]
  ];

  var form = document.getElementById("orderForm");
  var list = document.getElementById("orderList");
  var count = document.getElementById("count");
  var msg = document.getElementById("msg");
  var saveBtn = document.getElementById("saveBtn");
  var orders = load();

  function load() {
    try { return JSON.parse(localStorage.getItem(KEY)) || []; } catch (e) { return []; }
  }
  function persist() {
    try { localStorage.setItem(KEY, JSON.stringify(orders)); } catch (e) {
      say("Could not save on this device. Export your orders soon.", "err");
    }
  }
  function say(text, type) { msg.textContent = text; msg.className = type || ""; }

  function today() {
    var d = new Date();
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  }
  form.orderDate.value = today();

  function render() {
    list.innerHTML = "";
    count.textContent = orders.length;
    if (!orders.length) {
      var e = document.createElement("li");
      e.className = "empty";
      e.textContent = "No orders yet. Fill in the form above and save your first order.";
      e.style.border = "none";
      list.appendChild(e);
      return;
    }
    orders.slice().reverse().forEach(function (o) {
      var li = document.createElement("li");
      li.dataset.status = o.status;

      var info = document.createElement("div");
      info.className = "info";
      var n = document.createElement("strong");
      n.textContent = o.name;
      var d = document.createElement("span");
      d.textContent = o.design;
      info.append(n, d);

      var tools = document.createElement("div");
      tools.className = "row-tools";
      var sel = document.createElement("select");
      sel.setAttribute("aria-label", "Status for " + o.name);
      STATUSES.forEach(function (s) {
        var op = document.createElement("option");
        op.value = op.textContent = s;
        if (s === o.status) op.selected = true;
        sel.appendChild(op);
      });
      sel.addEventListener("change", function () {
        o.status = sel.value;
        li.dataset.status = o.status;
        persist();
      });
      var del = document.createElement("button");
      del.type = "button";
      del.className = "del";
      del.textContent = "Delete";
      del.setAttribute("aria-label", "Delete order for " + o.name);
      del.addEventListener("click", function () {
        if (confirm("Delete the order for " + o.name + "?")) {
          orders = orders.filter(function (x) { return x.id !== o.id; });
          persist(); render();
        }
      });
      tools.append(sel, del);
      li.append(info, tools);
      list.appendChild(li);
    });
  }

  form.addEventListener("submit", function (ev) {
    ev.preventDefault();
    var bad = false;
    ["name", "design"].forEach(function (f) {
      var el = form[f];
      var empty = !el.value.trim();
      el.classList.toggle("invalid", empty);
      if (empty) bad = true;
    });
    if (bad) { say("Please fill in the Name and Design / item fields.", "err"); return; }

    var order = { id: Date.now(), status: "New" };
    FIELDS.forEach(function (f) { order[f[0]] = form[f[0]].value.trim(); });
    orders.push(order);
    persist();
    render();

    var payload = {
      access_key: ACCESS_KEY,
      subject: "New WallCrafted order: " + order.name + " - " + order.design,
      from_name: "WallCrafted Order Log"
    };
    FIELDS.forEach(function (f) { payload[f[1]] = order[f[0]] || "-"; });
    if (order.email) payload.email = order.email; // lets you reply straight to the customer

    saveBtn.disabled = true;
    say("Order saved. Sending to email...", "");
    fetch("https://api.web3forms.com/submit", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Accept": "application/json" },
      body: JSON.stringify(payload)
    })
      .then(function (r) { return r.json(); })
      .then(function (res) {
        if (res.success) {
          say("Order saved and emailed to " + EMAIL + ".", "ok");
        } else {
          say("Order saved, but the email was not sent: " + (res.message || "unknown error"), "warn");
        }
      })
      .catch(function () {
        say("Order saved on this device, but the email could not be sent. Check your connection.", "warn");
      })
      .then(function () {
        saveBtn.disabled = false;
        form.reset();
        form.orderDate.value = today();
        form.quantity.value = 1;
        form.name.focus();
      });
  });

  form.addEventListener("input", function (e) { e.target.classList.remove("invalid"); });

  // Export to .xlsx (real Excel file). Every cell is stored as text so nothing gets changed.
  document.getElementById("exportBtn").addEventListener("click", function () {
    if (!orders.length) { say("There are no orders to export yet.", "warn"); return; }
    if (typeof XLSX === "undefined") {
      say("The Excel library did not load. Check your internet connection and reload the page.", "err");
      return;
    }
    var rows = [FIELDS.map(function (f) { return f[1]; }).concat("Status")];
    orders.forEach(function (o) {
      rows.push(FIELDS.map(function (f) { return String(o[f[0]] == null ? "" : o[f[0]]); }).concat(o.status));
    });
    var ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = rows[0].map(function (h, i) {
      var w = h.length;
      rows.forEach(function (r) { w = Math.max(w, String(r[i]).length); });
      return { wch: Math.min(w + 2, 40) };
    });
    var wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Orders");
    XLSX.writeFile(wb, "WallCrafted-orders-" + today() + ".xlsx");
  });

  render();
})();
