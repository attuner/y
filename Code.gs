/**
 * YAADY'S MILLET ROTI MEALS - MASTER SERVERLESS BACKEND (SECURED & HARDENED)
 * File: Code.gs
 * Version: 14.0.0 (Hardened Server-Side Calculations, Safe Whitelists, Concurrency & Tamper Resistance)
 */

const REFUND_RECEIPTS_FOLDER_NAME = "Yaadys_Wallet_Refund_Receipts";
const MENU_IMAGES_FOLDER_NAME = "Yaadys_Menu_Images";
const TIMEZONE_IST = "Asia/Kolkata";

// Whitelisted keys that can be exposed to public users
const PUBLIC_CONFIG_KEYS = [
  "KITCHEN_UPI_ID",
  "KITCHEN_NAME",
  "KITCHEN_PHONE",
  "KITCHEN_ADDRESS",
  "MIN_PREP_TIME_MINS",
  "WALLET_SYSTEM_ENABLED",
  "WALLET_SLAB_1_MIN",
  "WALLET_SLAB_1_DISC",
  "WALLET_SLAB_2_MIN",
  "WALLET_SLAB_2_DISC",
  "WALLET_SLAB_3_MIN",
  "WALLET_SLAB_3_DISC",
  "KITCHEN_OPEN",
  "AUTO_REFRESH_INTERVAL_SEC",
  "KITCHEN_OPEN_HOUR_IST",
  "KITCHEN_CLOSE_HOUR_IST",
  "ANNOUNCEMENT_BANNER"
];

// --------------------------------------------------------------------------
// TIME & DATE SANITIZATION HELPERS
// --------------------------------------------------------------------------
function cleanSheetDateString(val) {
  if (!val) return Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd");
  if (val instanceof Date) {
    return Utilities.formatDate(val, TIMEZONE_IST, "yyyy-MM-dd");
  }
  return String(val).split("T")[0].trim();
}

function cleanSheetTimeString(val) {
  if (!val) return "Flexible";
  if (val instanceof Date) {
    return Utilities.formatDate(val, TIMEZONE_IST, "hh:mm a");
  }
  const str = String(val).replace(/^'/, "").trim();
  if (str.includes("1899") || str.includes("GMT")) {
    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, TIMEZONE_IST, "hh:mm a");
    }
  }
  return str.replace(/\s*\(?IST\)?/gi, "").trim();
}

function getMasterAdminPin() {
  const propPin = PropertiesService.getScriptProperties().getProperty("ADMIN_PIN");
  if (propPin) return propPin;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Config");
  if (configSheet) {
    const data = configSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === "ADMIN_PIN") {
        return String(data[i][1]).trim();
      }
    }
  }
  return "1234";
}

function verifyAdminPin(providedPin) {
  if (!providedPin) return false;
  return String(providedPin).trim() === getMasterAdminPin();
}

// --------------------------------------------------------------------------
// REST API ROUTERS: doGet & doPost
// --------------------------------------------------------------------------
function doGet(e) {
  try {
    const params = e ? e.parameter : {};
    const action = params.action;

    switch (action) {
      case "getMenu":
        return sendJsonResponse(getMenuData());
      case "getCustomerProfile":
        return sendJsonResponse(getCustomerProfile(params.phone));
      case "getCustomerOrders":
        return sendJsonResponse(getCustomerOrders(params.phone));
      case "getOrderStatus":
        return sendJsonResponse(getOrderStatus(params.orderId));
      case "GET_ORDER_RECEIPT":
        return sendJsonResponse(getOrderReceiptData(params.orderId));
      case "getWalletLedger":
        return sendJsonResponse(getWalletLedger(params.customerId));
      case "getAdminData":
        return sendJsonResponse(getAdminData(params.pin));
      case "getConfig":
        return sendJsonResponse(getConfigData(false));
      case "trackVisit":
        return sendJsonResponse(trackVisitor(params.isRepeat === "true"));
      case "GET_CHAT_MESSAGES":
        return sendJsonResponse(getChatMessages(params.customerId));
      case "GET_ALL_CHAT_THREADS":
        return sendJsonResponse(getAllChatThreads(params.pin));
      case "GET_ONLINE_USERS":
        return sendJsonResponse(getOnlineUsers(params.pin));
      case "GET_DAYWISE_REPORT":
        return sendJsonResponse(getDaywiseReport(params.pin, params.date));
      case "GET_ANALYTICS_REPORT":
        return sendJsonResponse(getAnalyticsReport(params.pin));
      case "GET_HOURLY_TRAFFIC_REPORT":
        return sendJsonResponse(getHourlyTrafficReport(params.pin, params.date));
      default:
        return sendJsonResponse({
          status: "SUCCESS",
          message: "Yaady's Master API Online",
          timestamp: Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd'T'HH:mm:ssXXX")
        });
    }
  } catch (err) {
    return sendJsonResponse({ status: "ERROR", message: err.toString() });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (t) {
    return sendJsonResponse({ status: "ERROR", message: "Server busy processing concurrent requests. Please retry in a moment." });
  }

  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    const action = payload.action;

    switch (action) {
      case "REGISTER":
        return sendJsonResponse(registerCustomer(payload));
      case "LOGIN":
        return sendJsonResponse(loginCustomer(payload));
      case "FORGOT_PASSWORD":
        return sendJsonResponse(forgotPassword(payload));
      case "RESET_PASSWORD":
        return sendJsonResponse(resetPassword(payload));
      case "CHANGE_PASSWORD":
        return sendJsonResponse(changePassword(payload));
      case "CHECK_CART_VALIDITY":
        return sendJsonResponse(checkCartValidity(payload));
      case "INITIATE_PRE_ORDER":
        return sendJsonResponse(initiatePreOrder(payload));
      case "SUBMIT_FINAL_PAYMENT":
        return sendJsonResponse(submitFinalPayment(payload));
      case "SUBMIT_ORDER":
        return sendJsonResponse(submitOrder(payload));
      case "SUBMIT_WALLET_RECHARGE":
        return sendJsonResponse(submitWalletRecharge(payload));
      case "SUBMIT_WALLET_CLAIM":
        return sendJsonResponse(submitWalletClaim(payload));
      case "HEARTBEAT":
        return sendJsonResponse(recordHeartbeat(payload));
      case "SEND_CHAT_MESSAGE":
        return sendJsonResponse(sendChatMessage(payload));
      case "ADMIN_CONFIRM_PAYMENT":
        return sendJsonResponse(adminConfirmPayment(payload));
      case "ADMIN_SET_ORDER_STATUS":
        return sendJsonResponse(adminSetOrderStatus(payload));
      case "ADMIN_APPROVE_CLAIM":
        return sendJsonResponse(adminApproveClaim(payload));
      case "ADMIN_TOGGLE_MENU_STOCK":
        return sendJsonResponse(adminToggleMenuStock(payload));
      case "ADMIN_UPDATE_MENU_ITEM":
        return sendJsonResponse(adminUpdateMenuItem(payload));
      case "ADMIN_ADD_MENU_ITEM":
        return sendJsonResponse(adminAddMenuItem(payload));
      case "ADMIN_DELETE_MENU_ITEM":
        return sendJsonResponse(adminDeleteMenuItem(payload));
      case "ADMIN_UPDATE_CONFIG":
        return sendJsonResponse(adminUpdateConfig(payload));
      case "ADMIN_UPDATE_ANNOUNCEMENT":
        return sendJsonResponse(adminUpdateAnnouncement(payload));
      default:
        return sendJsonResponse({ status: "ERROR", message: "Unknown POST action: " + action });
    }
  } catch (err) {
    return sendJsonResponse({ status: "ERROR", message: err.toString() });
  } finally {
    lock.releaseLock();
  }
}

function sendJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------
// DATABASE SETUP & MIGRATION
// -------------------------------------------------------------
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sheets = [
    {
      name: "Customers",
      headers: ["Customer_ID", "Full_Name", "Mobile_Number", "Password_Hash", "WhatsApp_Number", "Email", "Wallet_Balance", "Total_Orders_Count", "Registered_At", "Account_Status", "Reset_Code", "Reset_Expires"]
    },
    {
      name: "Menu",
      headers: ["Item_ID", "Category", "Item_Name", "Description", "Millet_Type", "Base_Price", "Image_URL", "Is_Available", "Prep_Time_Mins", "Nutritional_Info"]
    },
    {
      name: "Orders",
      headers: ["Order_ID", "Timestamp", "Customer_ID", "Customer_Name", "Customer_Phone", "Items_JSON", "Total_Items_Count", "Subtotal", "Discount_Applied", "Discount_Slab", "Final_Payable", "Payment_Mode", "UTR_Number", "Payment_Screenshot_Drive_URL", "Pickup_Date", "Pickup_Time", "Order_Status", "Admin_Notes"]
    },
    {
      name: "Wallet_Ledger",
      headers: ["Txn_ID", "Timestamp", "Customer_ID", "Txn_Type", "Amount", "UTR_Ref", "Status", "Balance_After_Txn", "Drive_Receipt_URL", "Payout_UPI", "Withdrawal_Reason"]
    },
    {
      name: "Config",
      headers: ["Key", "Value", "Description"]
    },
    {
      name: "Analytics",
      headers: ["Date", "New_Visits", "Repeat_Visits", "Total_Visits", "New_Registrations", "Total_Orders", "Gross_Revenue"]
    },
    {
      name: "Hourly_Traffic",
      headers: ["Date", "Hour", "New_Visits", "Repeat_Visits", "Total_Visits", "Orders_Placed"]
    },
    {
      name: "Chat_Messages",
      headers: ["Message_ID", "Timestamp", "Customer_ID", "Customer_Name", "Sender", "Message_Text", "Is_Read"]
    },
    {
      name: "Active_Sessions",
      headers: ["Customer_ID", "Customer_Name", "Phone", "Last_Heartbeat"]
    }
  ];

  sheets.forEach(function(s) {
    let sheet = ss.getSheetByName(s.name);
    if (!sheet) {
      sheet = ss.insertSheet(s.name);
      sheet.appendRow(s.headers);
      sheet.getRange(1, 1, 1, s.headers.length).setFontWeight("bold").setBackground("#EDE8DE");
      sheet.setFrozenRows(1);
    } else {
      const currentHeaders = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
      s.headers.forEach((h) => {
        if (!currentHeaders.includes(h)) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h).setFontWeight("bold").setBackground("#EDE8DE");
        }
      });
    }
  });

  const configSheet = ss.getSheetByName("Config");
  if (configSheet.getLastRow() === 1) {
    const defaultConfigs = [
      ["KITCHEN_UPI_ID", "yaadyskitchen@upi", "UPI VPA address for customer payments"],
      ["KITCHEN_NAME", "Yaady's Millet Roti Meals", "Display Brand Name"],
      ["KITCHEN_PHONE", "+919000001690", "Primary Kitchen Helpline & WhatsApp"],
      ["NOTIFICATION_EMAIL", Session.getEffectiveUser().getEmail() || "admin@yaadys.com", "Admin Alert Email recipient"],
      ["KITCHEN_ADDRESS", "Neknampur, Manikonda, Hyderabad", "Pickup Counter Address"],
      ["MIN_PREP_TIME_MINS", "15", "Minimum baseline kitchen preparation minutes"],
      ["WALLET_SYSTEM_ENABLED", "TRUE", "Master switch for wallet prepaid perks and checkout"],
      ["WALLET_SLAB_1_MIN", "500", "Tier 1 Wallet Balance minimum"],
      ["WALLET_SLAB_1_DISC", "0.10", "Tier 1 Cart Discount (10%)"],
      ["WALLET_SLAB_2_MIN", "1000", "Tier 2 Wallet Balance minimum"],
      ["WALLET_SLAB_2_DISC", "0.15", "Tier 2 Cart Discount (15%)"],
      ["WALLET_SLAB_3_MIN", "1999", "Tier 3 Wallet Balance minimum"],
      ["WALLET_SLAB_3_DISC", "0.20", "Tier 3 Cart Discount (20%)"],
      ["KITCHEN_OPEN", "TRUE", "Kitchen order acceptance master toggle"],
      ["AUTO_REFRESH_INTERVAL_SEC", "10", "Client polling interval"],
      ["ADMIN_PIN", "1234", "Master Partner Portal Access PIN"],
      ["KITCHEN_OPEN_HOUR_IST", "10:00", "Kitchen opening time (HH:mm)"],
      ["KITCHEN_CLOSE_HOUR_IST", "22:00", "Kitchen closing time (HH:mm)"],
      ["ANNOUNCEMENT_BANNER", "Fresh Harvest Specials: Hot Jowar & Ragi Rotis made live upon pickup!", "Broadcast Banner"]
    ];
    defaultConfigs.forEach(function(c) { configSheet.appendRow(c); });
  }

  getOrCreateFolder(REFUND_RECEIPTS_FOLDER_NAME);
  getOrCreateFolder(MENU_IMAGES_FOLDER_NAME);
}

function getOrCreateFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) return folders.next();
  const f = DriveApp.createFolder(folderName);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return f;
}

function hashString(str) {
  const salt = "YaadysMilletKitchen2026SaltSecureV14";
  const rawBytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + String(str), Utilities.Charset.UTF_8);
  let hash = "";
  for (let i = 0; i < rawBytes.length; i++) {
    let byteVal = rawBytes[i];
    if (byteVal < 0) byteVal += 256;
    let hex = byteVal.toString(16);
    if (hex.length === 1) hex = "0" + hex;
    hash += hex;
  }
  return hash;
}

// -------------------------------------------------------------
// SERVER-SIDE BILL CALCULATION & PRICE RECONCILIATION
// -------------------------------------------------------------
function calculateServerBill(cartItems, customerWalletBalance) {
  const menuSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const menuData = menuSheet.getDataRange().getValues();
  const config = getConfigData(true).data;

  let subtotal = 0;
  let totalCount = 0;
  const verifiedItems = [];

  for (const cItem of cartItems) {
    const rawQty = parseInt(cItem.quantity, 10);
    if (isNaN(rawQty) || rawQty <= 0) continue;

    const itemId = String(cItem.id).trim();
    for (let i = 1; i < menuData.length; i++) {
      if (String(menuData[i][0]).trim() === itemId) {
        const itemPrice = Number(menuData[i][5]) || 0;
        const lineTotal = itemPrice * rawQty;
        subtotal += lineTotal;
        totalCount += rawQty;
        verifiedItems.push({
          id: itemId,
          name: String(menuData[i][2]),
          price: itemPrice,
          quantity: rawQty,
          lineTotal: lineTotal
        });
        break;
      }
    }
  }

  let discRate = 0;
  let slab = "0%";
  const walletEnabled = String(config.WALLET_SYSTEM_ENABLED).toUpperCase() === "TRUE";

  if (walletEnabled && customerWalletBalance !== undefined && customerWalletBalance !== null) {
    const bal = Number(customerWalletBalance) || 0;
    if (bal >= Number(config.WALLET_SLAB_3_MIN || 1999)) {
      discRate = Number(config.WALLET_SLAB_3_DISC || 0.20);
      slab = (discRate * 100) + "%";
    } else if (bal >= Number(config.WALLET_SLAB_2_MIN || 1000)) {
      discRate = Number(config.WALLET_SLAB_2_DISC || 0.15);
      slab = (discRate * 100) + "%";
    } else if (bal >= Number(config.WALLET_SLAB_1_MIN || 500)) {
      discRate = Number(config.WALLET_SLAB_1_DISC || 0.10);
      slab = (discRate * 100) + "%";
    }
  }

  const discountApplied = Math.round(subtotal * discRate);
  const finalPayable = Math.max(0, subtotal - discountApplied);

  return {
    subtotal: subtotal,
    discountApplied: discountApplied,
    discountSlab: slab,
    finalPayable: finalPayable,
    totalItemsCount: totalCount,
    items: verifiedItems
  };
}

// -------------------------------------------------------------
// INVENTORY & KITCHEN OPEN/CLOSED STRICT VALIDATION
// -------------------------------------------------------------
function checkCartValidity(payload) {
  const config = getConfigData(false).data;
  const isKitchenOpen = (String(config.KITCHEN_OPEN).toUpperCase() === "TRUE");

  if (!isKitchenOpen) {
    return {
      status: "KITCHEN_CLOSED",
      kitchenOpen: false,
      message: "The kitchen is currently closed for pickups."
    };
  }

  const menuSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const menuData = menuSheet.getDataRange().getValues();
  const cartItems = payload.cartItems || [];

  const unavailableItems = [];
  const validItems = [];

  cartItems.forEach(cartItem => {
    let matched = false;
    for (let i = 1; i < menuData.length; i++) {
      if (String(menuData[i][0]).trim() === String(cartItem.id).trim()) {
        matched = true;
        const available = (menuData[i][7] === true || String(menuData[i][7]).toUpperCase() === "TRUE");
        if (available) {
          validItems.push({
            id: menuData[i][0],
            name: menuData[i][2],
            price: Number(menuData[i][5]) || 0,
            quantity: Math.max(1, parseInt(cartItem.quantity, 10) || 1)
          });
        } else {
          unavailableItems.push({ id: menuData[i][0], name: menuData[i][2] });
        }
        break;
      }
    }
    if (!matched) {
      unavailableItems.push({ id: cartItem.id, name: "Discontinued Meal" });
    }
  });

  return {
    status: "SUCCESS",
    kitchenOpen: true,
    hasSoldOutItems: unavailableItems.length > 0,
    unavailableItems: unavailableItems,
    validItems: validItems
  };
}

// -------------------------------------------------------------
// STEP 1: PRE-ORDER INITIATION
// -------------------------------------------------------------
function initiatePreOrder(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  const customersSheet = ss.getSheetByName("Customers");
  const config = getConfigData(true).data;

  if (String(config.KITCHEN_OPEN).toUpperCase() !== "TRUE") {
    return { status: "ERROR", message: "Kitchen is currently closed for pickups." };
  }

  const orderId = String(payload.preAllocatedOrderId || ("ORD-" + Math.floor(100000 + Math.random() * 900000))).trim();
  const oData = ordersSheet.getDataRange().getValues();

  for (let i = 1; i < oData.length; i++) {
    if (String(oData[i][0]).trim() === orderId) {
      return { status: "SUCCESS", message: "Pre-order reservation confirmed.", orderId: orderId };
    }
  }

  let custBal = 0;
  const cData = customersSheet.getDataRange().getValues();
  for (let c = 1; c < cData.length; c++) {
    if (String(cData[c][0]).trim() === String(payload.customerId).trim()) {
      custBal = Number(cData[c][6]) || 0;
      break;
    }
  }

  const rawItems = payload.items || [];
  const serverBill = calculateServerBill(rawItems, custBal);

  if (serverBill.totalItemsCount === 0) {
    return { status: "ERROR", message: "Cannot place order with zero valid items." };
  }

  const nowIst = new Date();
  const istTimestamp = Utilities.formatDate(nowIst, TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");
  const cleanPickupDate = String(payload.pickupDate || "").split("T")[0].trim();
  const cleanPickupTime = String(payload.pickupTime || "").replace(/\s*\(?IST\)?/gi, "").trim();
  const storedPickupTimeString = cleanPickupTime ? "'" + cleanPickupTime : "'Flexible";

  ordersSheet.appendRow([
    orderId,
    istTimestamp,
    String(payload.customerId || "").trim(),
    String(payload.customerName || "Customer").trim(),
    String(payload.customerPhone || "").trim(),
    JSON.stringify(serverBill.items),
    serverBill.totalItemsCount,
    serverBill.subtotal,
    serverBill.discountApplied,
    serverBill.discountSlab,
    serverBill.finalPayable,
    "Pending_Payment",
    "Awaiting_Payment",
    "",
    cleanPickupDate,
    storedPickupTimeString,
    "Payment_Pending",
    "Early Pre-Order Reservation"
  ]);

  return {
    status: "SUCCESS",
    message: "Pre-order reservation recorded.",
    orderId: orderId,
    finalPayable: serverBill.finalPayable
  };
}

// -------------------------------------------------------------
// STEP 2: FINAL PAYMENT SUBMISSION & RECONCILIATION
// -------------------------------------------------------------
function submitFinalPayment(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  const customersSheet = ss.getSheetByName("Customers");
  const ledgerSheet = ss.getSheetByName("Wallet_Ledger");
  const config = getConfigData(true).data;

  const orderId = String(payload.orderId).trim();
  const oData = ordersSheet.getDataRange().getValues();
  let orderRow = -1;
  let orderRecord = null;

  for (let i = 1; i < oData.length; i++) {
    if (String(oData[i][0]).trim() === orderId) {
      orderRow = i + 1;
      orderRecord = oData[i];
      break;
    }
  }

  if (orderRow === -1) {
    return submitOrder(payload);
  }

  // Prevent double payment confirmation if already verified or accepted
  const currentStatus = String(orderRecord[16]).trim();
  if (currentStatus === "Kitchen_Accepted" || currentStatus === "Preparing" || currentStatus === "Ready_For_Pickup" || currentStatus === "Completed") {
    return { status: "SUCCESS", message: "Payment already verified.", orderId: orderId, orderStatus: currentStatus };
  }

  const payable = Number(orderRecord[10]) || 0;
  const customerId = String(orderRecord[2] || payload.customerId).trim();
  const customerPhone = String(orderRecord[4]).trim();
  const mode = String(payload.paymentMode || "Direct_UPI").trim();
  const utr = String(payload.utrNumber || "").trim();

  let targetStatus = "Pending_Verification";

  if (mode === "Wallet") {
    if (String(config.WALLET_SYSTEM_ENABLED).toUpperCase() !== "TRUE") {
      return { status: "ERROR", message: "Wallet payment paused. Please pay via Direct UPI." };
    }

    const cData = customersSheet.getDataRange().getValues();
    let custRow = -1;
    let currentBal = 0;

    for (let c = 1; c < cData.length; c++) {
      if (String(cData[c][0]).trim() === customerId || String(cData[c][2]).trim() === customerPhone) {
        custRow = c + 1;
        currentBal = Number(cData[c][6]) || 0;
        break;
      }
    }

    if (custRow === -1 || currentBal < payable) {
      return { status: "ERROR", message: "Insufficient wallet balance." };
    }

    const newBal = currentBal - payable;
    customersSheet.getRange(custRow, 7).setValue(newBal);

    const nowIst = new Date();
    const istTimestamp = Utilities.formatDate(nowIst, TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");
    const txnId = "TXN-" + Date.now();

    ledgerSheet.appendRow([
      txnId, istTimestamp, customerId, "Debit_Order", payable,
      "Order " + orderId, "Approved", newBal, "", "", "Meal Purchase"
    ]);

    targetStatus = "Kitchen_Accepted";
  } else {
    if (!/^\d{12}$/.test(utr)) {
      return { status: "ERROR", message: "UTR must be a valid 12-digit number." };
    }

    for (let k = 1; k < oData.length; k++) {
      if (k + 1 !== orderRow && String(oData[k][12]).trim() === utr) {
        return { status: "ERROR", message: "This UTR number has already been used for another order." };
      }
    }
  }

  ordersSheet.getRange(orderRow, 12).setValue(mode);
  ordersSheet.getRange(orderRow, 13).setValue(mode === "Wallet" ? "WALLET_DEDUCT" : utr);
  ordersSheet.getRange(orderRow, 17).setValue(targetStatus);

  recordOrderAnalytics(Number(orderRecord[6]) || 1, payable);
  recordHourlyOrderPlacement();

  try {
    sendAdminNotificationEmail("ORDER", {
      orderId: orderId,
      customerName: orderRecord[3],
      customerPhone: orderRecord[4],
      finalPayable: payable,
      items: JSON.parse(orderRecord[5] || "[]"),
      pickupDate: orderRecord[14],
      pickupTime: cleanSheetTimeString(orderRecord[15]),
      paymentMode: mode,
      utrNumber: utr || "N/A"
    });
  } catch (emailErr) {}

  return {
    status: "SUCCESS",
    message: "Payment submitted successfully!",
    orderId: orderId,
    orderStatus: targetStatus
  };
}

// -------------------------------------------------------------
// RECEIPT DATA PROVIDER
// -------------------------------------------------------------
function getOrderReceiptData(orderId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  const data = ordersSheet.getDataRange().getValues();
  const cleanId = String(orderId || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === cleanId) {
      return {
        status: "SUCCESS",
        order: {
          orderId: cleanId,
          token: cleanId.slice(-4),
          customerName: String(data[i][3]).trim(),
          customerPhone: String(data[i][4]).trim(),
          pickupDate: cleanSheetDateString(data[i][14]),
          pickupTime: cleanSheetTimeString(data[i][15]),
          paymentMode: String(data[i][11]).trim(),
          utrNumber: String(data[i][12] || "").trim(),
          orderStatus: String(data[i][16]).trim(),
          subtotal: Number(data[i][7]) || Number(data[i][10]) || 0,
          discountApplied: Number(data[i][8]) || 0,
          discountSlab: String(data[i][9] || ""),
          finalPayable: Number(data[i][10]) || 0,
          items: JSON.parse(data[i][5] || "[]")
        }
      };
    }
  }

  return { status: "ERROR", message: "Order reference not found." };
}

// -------------------------------------------------------------
// DIRECT FALLBACK ORDER SUBMISSION
// -------------------------------------------------------------
function submitOrder(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customersSheet = ss.getSheetByName("Customers");
  const ordersSheet = ss.getSheetByName("Orders");
  const ledgerSheet = ss.getSheetByName("Wallet_Ledger");
  const config = getConfigData(true).data;

  if (String(config.KITCHEN_OPEN).toUpperCase() !== "TRUE") {
    return { status: "ERROR", message: "The kitchen is currently closed for pickups. Orders cannot be scheduled." };
  }

  const phone = String(payload.customerPhone || "").trim();
  const custData = customersSheet.getDataRange().getValues();
  let custRowIndex = -1;
  let currentBalance = 0;
  let customerId = payload.customerId;

  for (let i = 1; i < custData.length; i++) {
    if (String(custData[i][2]).trim() === phone) {
      custRowIndex = i + 1;
      customerId = custData[i][0];
      currentBalance = Number(custData[i][6]) || 0;
      break;
    }
  }

  if (custRowIndex === -1) {
    return { status: "ERROR", message: "Customer account not found. Please log in again." };
  }

  const serverBill = calculateServerBill(payload.items || [], currentBalance);
  if (serverBill.totalItemsCount === 0) {
    return { status: "ERROR", message: "No valid items selected." };
  }

  const orderId = payload.preAllocatedOrderId || ("ORD-" + Math.floor(100000 + Math.random() * 900000));
  const payable = serverBill.finalPayable;
  let orderStatus = "Pending_Verification";

  const nowIst = new Date();
  const istFormattedTimestamp = Utilities.formatDate(nowIst, TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");

  if (payload.paymentMode === "Wallet") {
    if (currentBalance < payable) {
      return { status: "ERROR", message: "Insufficient wallet balance." };
    }
    const newBalance = currentBalance - payable;
    customersSheet.getRange(custRowIndex, 7).setValue(newBalance);

    const txnId = "TXN-" + Date.now();
    ledgerSheet.appendRow([
      txnId, istFormattedTimestamp, customerId, "Debit_Order", payable,
      "Order " + orderId, "Approved", newBalance, "", "", "Meal Purchase"
    ]);

    orderStatus = "Kitchen_Accepted";
  }

  const currentOrders = Number(customersSheet.getRange(custRowIndex, 8).getValue()) || 0;
  customersSheet.getRange(custRowIndex, 8).setValue(currentOrders + 1);

  const cleanPickupDate = String(payload.pickupDate || "").split("T")[0].trim();
  const cleanPickupTime = String(payload.pickupTime || "").replace(/\s*\(?IST\)?/gi, "").trim();
  const storedPickupTimeString = cleanPickupTime ? "'" + cleanPickupTime : "'Flexible";

  ordersSheet.appendRow([
    orderId,
    istFormattedTimestamp,
    customerId,
    String(payload.customerName || "Customer").trim(),
    phone,
    JSON.stringify(serverBill.items),
    serverBill.totalItemsCount,
    serverBill.subtotal,
    serverBill.discountApplied,
    serverBill.discountSlab,
    payable,
    String(payload.paymentMode || "Direct_UPI"),
    String(payload.utrNumber || (payload.paymentMode === "Wallet" ? "WALLET_DEDUCT" : "")),
    "",
    cleanPickupDate,
    storedPickupTimeString,
    orderStatus,
    String(payload.adminNotes || "")
  ]);

  recordOrderAnalytics(serverBill.totalItemsCount, payable);
  recordHourlyOrderPlacement();

  return {
    status: "SUCCESS",
    message: "Order placed successfully!",
    orderId: orderId,
    orderStatus: orderStatus
  };
}

function submitWalletRecharge(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customersSheet = ss.getSheetByName("Customers");
  const ledgerSheet = ss.getSheetByName("Wallet_Ledger");
  const config = getConfigData(true).data;

  if (String(config.WALLET_SYSTEM_ENABLED).toUpperCase() !== "TRUE") {
    return { status: "ERROR", message: "Wallet recharge is currently paused by admin." };
  }

  const phone = String(payload.phone || "").trim();
  const custData = customersSheet.getDataRange().getValues();
  let customerId = "";
  let currentBalance = 0;

  for (let i = 1; i < custData.length; i++) {
    if (String(custData[i][2]).trim() === phone) {
      customerId = custData[i][0];
      currentBalance = Number(custData[i][6]) || 0;
      break;
    }
  }

  if (!customerId) return { status: "ERROR", message: "Customer profile not found." };

  const txnId = String(payload.preAllocatedTxnId || ("WLT-" + Date.now())).trim();
  const amount = Number(payload.amount) || 0;
  const utr = String(payload.utrNumber || "").trim();

  if (amount <= 0) {
    return { status: "ERROR", message: "Invalid recharge amount." };
  }

  if (!/^\d{12}$/.test(utr)) {
    return { status: "ERROR", message: "UTR must be a valid 12-digit numerical reference." };
  }

  const ledgerData = ledgerSheet.getDataRange().getValues();
  for (let j = 1; j < ledgerData.length; j++) {
    if (String(ledgerData[j][5]).trim() === utr) {
      return { status: "ERROR", message: "This UTR number has already been submitted." };
    }
  }

  const istFormattedTimestamp = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");

  ledgerSheet.appendRow([
    txnId, istFormattedTimestamp, customerId, "Credit_Recharge", amount,
    utr, "Pending_Verification", currentBalance, "", "", `Wallet Pack Top-Up (${txnId})`
  ]);

  return {
    status: "SUCCESS",
    message: `Recharge of ₹${amount} submitted (Token #${txnId}). Funds will reflect once verified.`,
    txnId: txnId
  };
}

function sendAdminNotificationEmail(type, details) {
  const config = getConfigData(true).data;
  const adminEmail = config.NOTIFICATION_EMAIL || Session.getEffectiveUser().getEmail();
  if (!adminEmail || !adminEmail.includes("@")) return;

  let subject = "";
  let bodyHtml = "";

  if (type === "ORDER") {
    const token = String(details.orderId).slice(-4);
    subject = `🔔 Order Placed [Token #${token}] - ₹${details.finalPayable}`;
    const itemsHtml = (details.items || []).map(i => `<tr><td style="padding:8px 12px;border-bottom:1px solid #eee;">${i.name}</td><td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:center;">${i.quantity}</td><td style="padding:8px 12px;border-bottom:1px solid #eee;text-align:right;">₹${Number(i.lineTotal) || (Number(i.price) * Number(i.quantity))}</td></tr>`).join("");

    bodyHtml = `
      <div style="font-family:'Segoe UI',sans-serif;max-width:600px;margin:auto;border:1px solid #e0e0e0;border-radius:12px;overflow:hidden;">
        <div style="background:#1C3D2B;color:#fff;padding:20px;text-align:center;">
          <h2 style="margin:0;color:#F59E0B;">YAADY'S MILLET ROTI MEALS</h2>
          <p style="margin:5px 0 0;font-size:14px;color:#d0e0d5;">New Pickup Order Received</p>
        </div>
        <div style="padding:24px;">
          <div style="background:#FAF7F2;padding:16px;border-radius:8px;margin-bottom:20px;">
            <p style="margin:0;font-size:20px;font-weight:bold;color:#1C3D2B;">Pickup Token: #${token} (${details.orderId})</p>
            <p style="margin:6px 0 0;font-size:14px;color:#555;">Scheduled Pickup: <strong>${details.pickupDate} at ${details.pickupTime}</strong></p>
          </div>
          <table style="width:100%;font-size:14px;margin-bottom:20px;">
            <tr><td style="color:#666;">Customer Name:</td><td><strong>${details.customerName}</strong></td></tr>
            <tr><td style="color:#666;">Contact Phone:</td><td><strong>+91 ${details.customerPhone}</strong></td></tr>
            <tr><td style="color:#666;">Payment Mode:</td><td><strong>${details.paymentMode}</strong></td></tr>
            <tr><td style="color:#666;">Bank UTR / Ref:</td><td><strong>${details.utrNumber}</strong></td></tr>
          </table>
          <table style="width:100%;border-collapse:collapse;font-size:14px;margin-bottom:20px;">
            <thead>
              <tr style="background:#eee;"><th style="padding:8px 12px;text-align:left;">Item</th><th style="padding:8px 12px;text-align:center;">Qty</th><th style="padding:8px 12px;text-align:right;">Price</th></tr>
            </thead>
            <tbody>${itemsHtml}</tbody>
          </table>
          <div style="background:#1C3D2B;color:#fff;padding:14px 20px;border-radius:8px;display:flex;justify-content:space-between;align-items:center;">
            <span style="font-size:15px;font-weight:bold;">Final Payable Amount:</span>
            <span style="font-size:22px;font-weight:bold;color:#F59E0B;">₹${details.finalPayable}</span>
          </div>
        </div>
      </div>
    `;
  }

  GmailApp.sendEmail(adminEmail, subject, "", { htmlBody: bodyHtml });
}

function trackVisitor(isRepeat) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analyticsSheet = ss.getSheetByName("Analytics");
  const hourlySheet = ss.getSheetByName("Hourly_Traffic");

  const nowIst = new Date();
  const todayStr = Utilities.formatDate(nowIst, TIMEZONE_IST, "yyyy-MM-dd");
  const currentHour = parseInt(Utilities.formatDate(nowIst, TIMEZONE_IST, "H"), 10);

  const aData = analyticsSheet.getDataRange().getValues();
  let dailyFound = false;
  for (let i = 1; i < aData.length; i++) {
    if (String(aData[i][0]) === todayStr) {
      let newVisits = Number(aData[i][1]) || 0;
      let repeatVisits = Number(aData[i][2]) || 0;
      let totalVisits = Number(aData[i][3]) || 0;

      if (isRepeat) repeatVisits += 1;
      else newVisits += 1;
      totalVisits += 1;

      analyticsSheet.getRange(i + 1, 2).setValue(newVisits);
      analyticsSheet.getRange(i + 1, 3).setValue(repeatVisits);
      analyticsSheet.getRange(i + 1, 4).setValue(totalVisits);
      dailyFound = true;
      break;
    }
  }
  if (!dailyFound) {
    analyticsSheet.appendRow([todayStr, isRepeat ? 0 : 1, isRepeat ? 1 : 0, 1, 0, 0, 0]);
  }

  const hData = hourlySheet.getDataRange().getValues();
  let hourlyFound = false;
  for (let j = 1; j < hData.length; j++) {
    if (String(hData[j][0]) === todayStr && Number(hData[j][1]) === currentHour) {
      let hNew = Number(hData[j][2]) || 0;
      let hRep = Number(hData[j][3]) || 0;
      let hTot = Number(hData[j][4]) || 0;

      if (isRepeat) hRep += 1;
      else hNew += 1;
      hTot += 1;

      hourlySheet.getRange(j + 1, 3).setValue(hNew);
      hourlySheet.getRange(j + 1, 4).setValue(hRep);
      hourlySheet.getRange(j + 1, 5).setValue(hTot);
      hourlyFound = true;
      break;
    }
  }
  if (!hourlyFound) {
    hourlySheet.appendRow([todayStr, currentHour, isRepeat ? 0 : 1, isRepeat ? 1 : 0, 1, 0]);
  }

  return { status: "SUCCESS", date: todayStr, hour: currentHour, isRepeat: isRepeat };
}

function recordHourlyOrderPlacement() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hourlySheet = ss.getSheetByName("Hourly_Traffic");
  const nowIst = new Date();
  const todayStr = Utilities.formatDate(nowIst, TIMEZONE_IST, "yyyy-MM-dd");
  const currentHour = parseInt(Utilities.formatDate(nowIst, TIMEZONE_IST, "H"), 10);

  const hData = hourlySheet.getDataRange().getValues();
  for (let j = 1; j < hData.length; j++) {
    if (String(hData[j][0]) === todayStr && Number(hData[j][1]) === currentHour) {
      const currentOrders = Number(hData[j][5]) || 0;
      hourlySheet.getRange(j + 1, 6).setValue(currentOrders + 1);
      return;
    }
  }
  hourlySheet.appendRow([todayStr, currentHour, 0, 0, 0, 1]);
}

function getHourlyTrafficReport(pin, targetDateStr) {
  if (!verifyAdminPin(pin)) {
    return { status: "ERROR", message: "Unauthorized PIN." };
  }

  const queryDate = targetDateStr || Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hourlySheet = ss.getSheetByName("Hourly_Traffic");
  const hData = hourlySheet.getDataRange().getValues();

  const hourlyMap = {};
  for (let h = 0; h < 24; h++) {
    const startHour12 = (h % 12 === 0 ? 12 : h % 12);
    const startPeriod = h < 12 ? "AM" : "PM";
    const nextH = (h + 1) % 24;
    const endHour12 = (nextH % 12 === 0 ? 12 : nextH % 12);
    const endPeriod = (h + 1) < 12 || (h + 1) === 24 ? "AM" : "PM";

    const label = `${String(startHour12).padStart(2, '0')}:00 ${startPeriod} - ${String(endHour12).padStart(2, '0')}:00 ${endPeriod}`;

    hourlyMap[h] = {
      hour: h,
      label: label,
      newVisits: 0,
      repeatVisits: 0,
      totalVisits: 0,
      ordersPlaced: 0
    };
  }

  for (let i = 1; i < hData.length; i++) {
    if (String(hData[i][0]) === queryDate) {
      const h = Number(hData[i][1]);
      if (hourlyMap[h]) {
        hourlyMap[h].newVisits += Number(hData[i][2]) || 0;
        hourlyMap[h].repeatVisits += Number(hData[i][3]) || 0;
        hourlyMap[h].totalVisits += Number(hData[i][4]) || 0;
        hourlyMap[h].ordersPlaced += Number(hData[i][5]) || 0;
      }
    }
  }

  return {
    status: "SUCCESS",
    date: queryDate,
    hourlyData: Object.values(hourlyMap)
  };
}

function getDaywiseReport(pin, targetDateStr) {
  if (!verifyAdminPin(pin)) return { status: "ERROR", message: "Unauthorized PIN." };

  const queryDate = targetDateStr || Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd");
  const ordersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  const data = ordersSheet.getDataRange().getValues();

  let totalOrders = 0, accepted = 0, preparing = 0, ready = 0, completed = 0, cancelled = 0, grossRevenue = 0, upiRevenue = 0, walletRevenue = 0, totalRotis = 0;
  const filteredOrders = [];

  for (let i = 1; i < data.length; i++) {
    const rawDate = data[i][14];
    const orderDate = cleanSheetDateString(rawDate);

    if (orderDate === queryDate) {
      const status = String(data[i][16]).trim();
      const amount = Number(data[i][10]) || 0;
      const mode = String(data[i][11]).trim();
      const itemsCount = Number(data[i][6]) || 0;

      totalOrders++;
      if (status === "Kitchen_Accepted") accepted++;
      else if (status === "Preparing") preparing++;
      else if (status === "Ready_For_Pickup") ready++;
      else if (status === "Completed") completed++;
      else if (status === "Rejected" || status === "Cancelled") cancelled++;

      if (status !== "Rejected" && status !== "Cancelled") {
        grossRevenue += amount;
        if (mode === "Wallet") walletRevenue += amount;
        else upiRevenue += amount;
        totalRotis += itemsCount;
      }

      filteredOrders.push({
        orderId: String(data[i][0]).trim(),
        timestamp: data[i][1],
        customerName: String(data[i][3]).trim(),
        customerPhone: String(data[i][4]).trim(),
        items: JSON.parse(data[i][5] || "[]"),
        itemsCount: itemsCount,
        finalPayable: amount,
        paymentMode: mode,
        utr: String(data[i][12] || "").trim(),
        pickupDate: orderDate,
        pickupTime: cleanSheetTimeString(data[i][15]),
        orderStatus: status
      });
    }
  }

  return {
    status: "SUCCESS",
    date: queryDate,
    metrics: { totalOrders, accepted, preparing, ready, completed, cancelled, grossRevenue, upiRevenue, walletRevenue, rotisPrepared: totalRotis },
    orders: filteredOrders
  };
}

function getCustomerOrders(phone) {
  const cleanPhone = String(phone || "").replace(/\D/g, '').slice(-10);
  if (!cleanPhone || cleanPhone.length !== 10) {
    return { status: "ERROR", message: "Invalid phone number." };
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  const data = sheet.getDataRange().getValues();
  const orders = [];

  for (let i = 1; i < data.length; i++) {
    const rowPhone = String(data[i][4] || "").replace(/\D/g, '').slice(-10);
    if (rowPhone === cleanPhone) {
      const cleanDate = cleanSheetDateString(data[i][14]);
      const cleanTime = cleanSheetTimeString(data[i][15]);

      orders.push({
        orderId: String(data[i][0]).trim(),
        timestamp: data[i][1],
        customerId: String(data[i][2]).trim(),
        customerName: String(data[i][3]).trim(),
        customerPhone: String(data[i][4]).trim(),
        items: JSON.parse(data[i][5] || "[]"),
        totalItemsCount: Number(data[i][6]) || 0,
        subtotal: Number(data[i][7]) || 0,
        discountApplied: Number(data[i][8]) || 0,
        discountSlab: String(data[i][9] || ""),
        finalPayable: Number(data[i][10]) || 0,
        paymentMode: String(data[i][11] || ""),
        utrNumber: String(data[i][12] || ""),
        pickupDate: cleanDate,
        pickupTime: cleanTime,
        orderStatus: String(data[i][16] || "").trim()
      });
    }
  }
  return { status: "SUCCESS", orders: orders.reverse() };
}

function getOrderStatus(orderId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  const data = sheet.getDataRange().getValues();
  const cleanOrderId = String(orderId || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === cleanOrderId) {
      const cleanDate = cleanSheetDateString(data[i][14]);
      const cleanTime = cleanSheetTimeString(data[i][15]);

      return {
        status: "SUCCESS",
        order: {
          orderId: String(data[i][0]).trim(),
          timestamp: data[i][1],
          customerName: String(data[i][3]).trim(),
          customerPhone: String(data[i][4]).trim(),
          items: JSON.parse(data[i][5] || "[]"),
          finalPayable: Number(data[i][10]) || 0,
          paymentMode: String(data[i][11] || ""),
          utrNumber: String(data[i][12] || ""),
          pickupDate: cleanDate,
          pickupTime: cleanTime,
          orderStatus: String(data[i][16] || "").trim()
        }
      };
    }
  }
  return { status: "ERROR", message: "Order not found." };
}

function getAnalyticsReport(pin) {
  if (!verifyAdminPin(pin)) {
    return { status: "ERROR", message: "Unauthorized PIN." };
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const analyticsSheet = ss.getSheetByName("Analytics");
  const aData = analyticsSheet.getDataRange().getValues();
  const customersSheet = ss.getSheetByName("Customers");
  const totalRegisteredUsers = Math.max(0, customersSheet.getLastRow() - 1);

  const dayWise = [];
  const weekWiseMap = {};

  for (let i = 1; i < aData.length; i++) {
    const dateStr = String(aData[i][0]);
    const newV = Number(aData[i][1]) || 0;
    const repV = Number(aData[i][2]) || 0;
    const totV = Number(aData[i][3]) || (newV + repV);
    const newReg = Number(aData[i][4]) || 0;
    const totOrders = Number(aData[i][5]) || 0;
    const grossRev = Number(aData[i][6]) || 0;

    dayWise.push({
      date: dateStr,
      newVisits: newV,
      repeatVisits: repV,
      totalVisits: totV,
      newRegistrations: newReg,
      orders: totOrders,
      revenue: grossRev
    });

    const d = new Date(dateStr + "T00:00:00+05:30");
    if (!isNaN(d.getTime())) {
      const startOfYear = new Date(d.getFullYear(), 0, 1);
      const weekNum = Math.ceil((((d - startOfYear) / 86400000) + startOfYear.getDay() + 1) / 7);
      const weekKey = `${d.getFullYear()} - W${weekNum}`;

      if (!weekWiseMap[weekKey]) {
        weekWiseMap[weekKey] = {
          week: weekKey,
          newVisits: 0,
          repeatVisits: 0,
          totalVisits: 0,
          newRegistrations: 0,
          orders: 0,
          revenue: 0
        };
      }
      weekWiseMap[weekKey].newVisits += newV;
      weekWiseMap[weekKey].repeatVisits += repV;
      weekWiseMap[weekKey].totalVisits += totV;
      weekWiseMap[weekKey].newRegistrations += newReg;
      weekWiseMap[weekKey].orders += totOrders;
      weekWiseMap[weekKey].revenue += grossRev;
    }
  }

  return {
    status: "SUCCESS",
    totalRegisteredUsers: totalRegisteredUsers,
    dayWise: dayWise.reverse(),
    weekWise: Object.values(weekWiseMap).reverse()
  };
}

function registerCustomer(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();
  const phone = String(payload.phone || "").replace(/\D/g, '').slice(-10);
  const password = String(payload.password || "").trim();

  if (!/^[6-9]\d{9}$/.test(phone)) return { status: "ERROR", message: "Please enter a valid 10-digit Indian Mobile Number." };
  if (!password || password.length < 4) return { status: "ERROR", message: "Password must be at least 4 characters long." };

  for (let i = 1; i < data.length; i++) {
    const existingPhone = String(data[i][2]).replace(/\D/g, '').slice(-10);
    if (existingPhone === phone) return { status: "ERROR", message: "Account already exists with this mobile number. Please log in." };
  }

  const randomHex = Math.floor(Math.random() * 0xFFFFFF).toString(16).padEnd(6, '0').toUpperCase();
  const customerId = "YMR-" + phone.slice(-4) + "-" + randomHex;
  const istTimestamp = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");

  sheet.appendRow([
    customerId, String(payload.fullName || "Guest").trim(), phone, hashString(password),
    phone, String(payload.email || "").trim(), 0, 0, istTimestamp, "Active", "", ""
  ]);

  recordRegistrationAnalytics();

  return {
    status: "SUCCESS",
    message: "Registration completed successfully.",
    customer: { customerId, fullName: payload.fullName, phone, whatsapp: phone, email: payload.email || "", walletBalance: 0, totalOrders: 0 }
  };
}

function recordRegistrationAnalytics() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analytics");
  const todayStr = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === todayStr) {
      sheet.getRange(i + 1, 5).setValue((Number(data[i][4]) || 0) + 1);
      return;
    }
  }
  sheet.appendRow([todayStr, 0, 0, 0, 1, 0, 0]);
}

function recordOrderAnalytics(rotisCount, amount) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Analytics");
  const todayStr = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === todayStr) {
      sheet.getRange(i + 1, 6).setValue((Number(data[i][5]) || 0) + 1);
      sheet.getRange(i + 1, 7).setValue((Number(data[i][6]) || 0) + amount);
      return;
    }
  }
  sheet.appendRow([todayStr, 0, 0, 0, 0, 1, amount]);
}

function loginCustomer(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();
  const phone = String(payload.phone || "").replace(/\D/g, '').slice(-10);
  const inputHash = hashString(String(payload.password || "").trim());

  for (let i = 1; i < data.length; i++) {
    const existingPhone = String(data[i][2]).replace(/\D/g, '').slice(-10);
    if (existingPhone === phone) {
      if (String(data[i][3]) === inputHash) {
        return {
          status: "SUCCESS",
          message: "Welcome back!",
          customer: { customerId: data[i][0], fullName: data[i][1], phone: data[i][2], whatsapp: data[i][4], email: data[i][5], walletBalance: Number(data[i][6]) || 0, totalOrders: Number(data[i][7]) || 0 }
        };
      } else {
        return { status: "ERROR", message: "Incorrect password. Please verify and try again." };
      }
    }
  }
  return { status: "ERROR", message: "Mobile number not registered. Please create an account." };
}

function getCustomerProfile(phone) {
  const cleanPhone = String(phone || "").replace(/\D/g, '').slice(-10);
  if (!cleanPhone || cleanPhone.length !== 10) return { status: "ERROR", message: "Invalid phone number." };

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    const rowPhone = String(data[i][2]).replace(/\D/g, '').slice(-10);
    if (rowPhone === cleanPhone) {
      return {
        status: "SUCCESS",
        customer: { customerId: data[i][0], fullName: data[i][1], phone: data[i][2], whatsapp: data[i][4], email: data[i][5], walletBalance: Number(data[i][6]) || 0, totalOrders: Number(data[i][7]) || 0, accountStatus: data[i][9] }
      };
    }
  }
  return { status: "ERROR", message: "Customer profile not found." };
}

function getMenuData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const data = sheet.getDataRange().getValues();
  const menu = [];

  for (let i = 1; i < data.length; i++) {
    menu.push({
      id: data[i][0],
      category: data[i][1],
      name: data[i][2],
      description: data[i][3],
      milletType: data[i][4],
      price: Number(data[i][5]) || 0,
      image: data[i][6],
      available: (data[i][7] === true || String(data[i][7]).toUpperCase() === "TRUE"),
      prepTimeMins: Number(data[i][8]) || 15,
      nutrition: data[i][9]
    });
  }
  return { status: "SUCCESS", menu: menu };
}

function getConfigData(includeInternal) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  const data = sheet.getDataRange().getValues();
  const config = {};
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][0]).trim();
    if (includeInternal || PUBLIC_CONFIG_KEYS.includes(k)) {
      config[k] = data[i][1];
    }
  }
  return { status: "SUCCESS", data: config };
}

function submitWalletClaim(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customersSheet = ss.getSheetByName("Customers");
  const ledgerSheet = ss.getSheetByName("Wallet_Ledger");

  const phone = String(payload.phone || "").replace(/\D/g, '').slice(-10);
  const claimAmount = Number(payload.amount) || 0;
  const payoutUpi = String(payload.payoutUpi || "").trim();
  const reason = String(payload.reason || "Customer refund request").trim();

  if (!payoutUpi.includes("@")) return { status: "ERROR", message: "Please provide a valid UPI ID (e.g., name@okaxis)." };

  const custData = customersSheet.getDataRange().getValues();
  let custRowIndex = -1;
  let customerId = "";
  let currentBalance = 0;

  for (let i = 1; i < custData.length; i++) {
    const rowPhone = String(custData[i][2]).replace(/\D/g, '').slice(-10);
    if (rowPhone === phone) {
      custRowIndex = i + 1;
      customerId = custData[i][0];
      currentBalance = Number(custData[i][6]) || 0;
      break;
    }
  }

  if (custRowIndex === -1) return { status: "ERROR", message: "Customer not found." };
  if (claimAmount <= 0 || claimAmount > currentBalance) {
    return { status: "ERROR", message: "Claim amount cannot exceed balance (₹" + currentBalance + ")." };
  }

  const remainingBal = currentBalance - claimAmount;
  customersSheet.getRange(custRowIndex, 7).setValue(remainingBal);

  const claimId = "CLM-" + Math.floor(10000 + Math.random() * 90000);
  const istTimestamp = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");

  ledgerSheet.appendRow([claimId, istTimestamp, customerId, "Debit_Claim", claimAmount, "HOLD", "CLAIM_PENDING", remainingBal, "", payoutUpi, reason]);

  return { status: "SUCCESS", message: "Claim logged. Payout to " + payoutUpi + " will be reviewed shortly.", claimId: claimId, newBalance: remainingBal };
}

function getWalletLedger(customerId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Wallet_Ledger");
  const data = sheet.getDataRange().getValues();
  const txns = [];
  const cleanId = String(customerId || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === cleanId) {
      txns.push({
        txnId: data[i][0],
        timestamp: data[i][1],
        type: data[i][3],
        amount: Number(data[i][4]) || 0,
        utrRef: data[i][5],
        status: data[i][6],
        balanceAfter: Number(data[i][7]) || 0,
        receiptPdfUrl: data[i][8],
        payoutUpi: data[i][9],
        reason: data[i][10]
      });
    }
  }
  return { status: "SUCCESS", transactions: txns.reverse() };
}

function getAdminData(pin) {
  if (!verifyAdminPin(pin)) return { status: "ERROR", message: "Invalid Kitchen Partner PIN." };

  const config = getConfigData(true).data;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders");
  const oData = ordersSheet.getDataRange().getValues();
  const orders = [];

  for (let i = 1; i < oData.length; i++) {
    const cleanDate = cleanSheetDateString(oData[i][14]);
    const cleanTime = cleanSheetTimeString(oData[i][15]);

    orders.push({
      orderId: String(oData[i][0]).trim(),
      timestamp: oData[i][1],
      customerId: String(oData[i][2]).trim(),
      customerName: String(oData[i][3]).trim(),
      customerPhone: String(oData[i][4]).trim(),
      items: JSON.parse(oData[i][5] || "[]"),
      totalItems: Number(oData[i][6]) || 0,
      finalPayable: Number(oData[i][10]) || 0,
      paymentMode: String(oData[i][11] || ""),
      utr: String(oData[i][12] || "").trim(),
      screenshotUrl: String(oData[i][13] || ""),
      pickupDate: cleanDate,
      pickupTime: cleanTime,
      orderStatus: String(oData[i][16] || "").trim()
    });
  }

  const ledgerSheet = ss.getSheetByName("Wallet_Ledger");
  const lData = ledgerSheet.getDataRange().getValues();
  const pendingRecharges = [];
  const pendingClaims = [];

  for (let j = 1; j < lData.length; j++) {
    if (lData[j][6] === "Pending_Verification") {
      pendingRecharges.push({
        txnId: lData[j][0], timestamp: lData[j][1],
        customerId: lData[j][2], amount: Number(lData[j][4]) || 0, utr: lData[j][5]
      });
    } else if (lData[j][6] === "CLAIM_PENDING") {
      pendingClaims.push({
        claimId: lData[j][0], timestamp: lData[j][1],
        customerId: lData[j][2], amount: Number(lData[j][4]) || 0, payoutUpi: lData[j][9], reason: lData[j][10]
      });
    }
  }

  return {
    status: "SUCCESS",
    config: config,
    orders: orders.reverse(),
    pendingRecharges: pendingRecharges,
    pendingClaims: pendingClaims
  };
}

function adminConfirmPayment(payload) {
  if (!verifyAdminPin(payload.pin)) return { status: "ERROR", message: "Unauthorized PIN." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (payload.type === "RECHARGE") {
    const ledgerSheet = ss.getSheetByName("Wallet_Ledger");
    const customersSheet = ss.getSheetByName("Customers");
    const txnId = String(payload.txnId || "").trim();
    const lData = ledgerSheet.getDataRange().getValues();

    let targetRow = -1;
    let customerId = "";
    let amount = 0;
    let currentStatus = "";

    for (let r = 1; r < lData.length; r++) {
      if (String(lData[r][0]).trim() === txnId) {
        targetRow = r + 1;
        customerId = String(lData[r][2]).trim();
        amount = Number(lData[r][4]) || 0;
        currentStatus = String(lData[r][6]).trim();
        break;
      }
    }

    if (targetRow === -1) return { status: "ERROR", message: "Transaction reference not found." };
    if (currentStatus === "Approved") return { status: "SUCCESS", message: "Recharge was already approved." };

    const custData = customersSheet.getDataRange().getValues();
    for (let i = 1; i < custData.length; i++) {
      if (String(custData[i][0]).trim() === customerId) {
        const newBal = (Number(custData[i][6]) || 0) + amount;
        customersSheet.getRange(i + 1, 7).setValue(newBal);
        ledgerSheet.getRange(targetRow, 7).setValue("Approved");
        ledgerSheet.getRange(targetRow, 8).setValue(newBal);
        return { status: "SUCCESS", message: "Wallet recharge credited & approved." };
      }
    }
    return { status: "ERROR", message: "Customer ID not matched." };
  } else {
    const ordersSheet = ss.getSheetByName("Orders");
    const orderId = String(payload.orderId || "").trim();
    const oData = ordersSheet.getDataRange().getValues();

    for (let r = 1; r < oData.length; r++) {
      if (String(oData[r][0]).trim() === orderId) {
        ordersSheet.getRange(r + 1, 17).setValue("Kitchen_Accepted");
        return { status: "SUCCESS", message: "Payment verified. Order accepted." };
      }
    }
    return { status: "ERROR", message: "Order not found." };
  }
}

function adminSetOrderStatus(payload) {
  if (!verifyAdminPin(payload.pin)) return { status: "ERROR", message: "Unauthorized PIN." };
  const ordersSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Orders");
  const orderId = String(payload.orderId || "").trim();
  const oData = ordersSheet.getDataRange().getValues();

  for (let r = 1; r < oData.length; r++) {
    if (String(oData[r][0]).trim() === orderId) {
      ordersSheet.getRange(r + 1, 17).setValue(payload.newStatus);
      return { status: "SUCCESS", message: "Order updated to " + payload.newStatus };
    }
  }
  return { status: "ERROR", message: "Order not found." };
}

function adminApproveClaim(payload) {
  if (!verifyAdminPin(payload.pin)) return { status: "ERROR", message: "Unauthorized PIN." };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledgerSheet = ss.getSheetByName("Wallet_Ledger");
  const claimId = String(payload.claimId || "").trim();
  const transferUtr = String(payload.transferUtr || "").trim();

  if (!/^\d{12}$/.test(transferUtr)) return { status: "ERROR", message: "Please provide a valid 12-digit bank transfer UTR." };

  const lData = ledgerSheet.getDataRange().getValues();
  for (let r = 1; r < lData.length; r++) {
    if (String(lData[r][0]).trim() === claimId) {
      ledgerSheet.getRange(r + 1, 6).setValue(transferUtr);
      ledgerSheet.getRange(r + 1, 7).setValue("Approved");
      return { status: "SUCCESS", message: "Claim disbursed successfully." };
    }
  }
  return { status: "ERROR", message: "Claim not found." };
}

function adminAddMenuItem(payload) {
  if (!verifyAdminPin(payload.pin)) return { status: "ERROR", message: "Unauthorized PIN." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const itemId = "ITEM-" + Math.floor(100 + Math.random() * 900);

  let imageUrl = "https://images.unsplash.com/photo-1546833999-b9f581a1996d?w=600";
  if (payload.base64Image && payload.base64Image.length > 50) {
    try {
      const folder = getOrCreateFolder(MENU_IMAGES_FOLDER_NAME);
      const mimeMatches = payload.base64Image.match(/^data:(image\/[a-zA-Z+]+);base64,/);
      const contentType = mimeMatches ? mimeMatches[1] : "image/jpeg";
      
      const allowedMimes = ["image/jpeg", "image/png", "image/webp"];
      if (allowedMimes.includes(contentType)) {
        const rawBase64 = payload.base64Image.substring(payload.base64Image.indexOf(",") + 1);
        const ext = contentType.split("/")[1].replace("jpeg", "jpg");
        const blob = Utilities.newBlob(Utilities.base64Decode(rawBase64), contentType, "Menu_" + itemId + "." + ext);
        const file = folder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        imageUrl = file.getUrl();
      }
    } catch (e) {}
  }

  sheet.appendRow([
    itemId, String(payload.category || "Specials").trim(), String(payload.name || "Untitled Meal").trim(), String(payload.description || "").trim(),
    String(payload.milletType || "Millet").trim(), Math.max(0, Number(payload.price) || 0), imageUrl, true,
    Math.max(5, Number(payload.prepTimeMins) || 15), String(payload.nutrition || "Balanced Fiber & Protein").trim()
  ]);

  return { status: "SUCCESS", message: "Meal added to active menu.", itemId: itemId };
}

function adminDeleteMenuItem(payload) {
  if (!verifyAdminPin(payload.pin)) return { status: "ERROR", message: "Unauthorized PIN." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.itemId).trim()) {
      sheet.deleteRow(i + 1);
      return { status: "SUCCESS", message: "Item removed from menu permanently." };
    }
  }
  return { status: "ERROR", message: "Item not found." };
}

function adminToggleMenuStock(payload) {
  if (!verifyAdminPin(payload.pin)) return { status: "ERROR", message: "Unauthorized PIN." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.itemId).trim()) {
      sheet.getRange(i + 1, 8).setValue(payload.available === true);
      return { status: "SUCCESS", message: "Stock status updated." };
    }
  }
  return { status: "ERROR", message: "Item not found." };
}

function adminUpdateMenuItem(payload) {
  if (!verifyAdminPin(payload.pin)) return { status: "ERROR", message: "Unauthorized PIN." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Menu");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(payload.itemId).trim()) {
      if (payload.price) sheet.getRange(i + 1, 6).setValue(Math.max(0, Number(payload.price) || 0));
      return { status: "SUCCESS", message: "Price updated." };
    }
  }
  return { status: "ERROR", message: "Item not found." };
}

function adminUpdateConfig(payload) {
  if (!verifyAdminPin(payload.pin)) return { status: "ERROR", message: "Unauthorized PIN." };
  
  const key = String(payload.key).trim();
  if (key === "ADMIN_PIN") {
    PropertiesService.getScriptProperties().setProperty("ADMIN_PIN", String(payload.value).trim());
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(String(payload.value));
      return { status: "SUCCESS", message: "Config key " + key + " set to " + payload.value };
    }
  }
  sheet.appendRow([key, payload.value, "Custom Configuration"]);
  return { status: "SUCCESS", message: "Config saved." };
}

function adminUpdateAnnouncement(payload) {
  if (!verifyAdminPin(payload.pin)) return { status: "ERROR", message: "Unauthorized PIN." };
  return adminUpdateConfig({
    pin: payload.pin,
    key: "ANNOUNCEMENT_BANNER",
    value: String(payload.bannerText || "").trim()
  });
}

function recordHeartbeat(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Active_Sessions");
  const now = new Date();
  const nowStr = Utilities.formatDate(now, TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");
  const custId = String(payload.customerId || "").trim();

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === custId) {
      sheet.getRange(i + 1, 4).setValue(nowStr);
      return { status: "SUCCESS" };
    }
  }
  sheet.appendRow([custId, payload.fullName || "Guest", payload.phone || "", nowStr]);
  return { status: "SUCCESS" };
}

function getOnlineUsers(pin) {
  if (!verifyAdminPin(pin)) return { status: "ERROR", message: "Unauthorized PIN." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Active_Sessions");
  const data = sheet.getDataRange().getValues();
  const nowMs = new Date().getTime();
  const activeUsers = [];

  for (let i = 1; i < data.length; i++) {
    const lastHb = new Date(data[i][3]).getTime();
    if (nowMs - lastHb <= 75000) {
      activeUsers.push({
        customerId: data[i][0],
        customerName: data[i][1],
        phone: data[i][2]
      });
    }
  }
  return { status: "SUCCESS", count: activeUsers.length, activeUsers: activeUsers };
}

function sendChatMessage(payload) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chat_Messages");
  const msgId = "MSG-" + Date.now();
  const istTimestamp = Utilities.formatDate(new Date(), TIMEZONE_IST, "yyyy-MM-dd HH:mm:ss");

  sheet.appendRow([
    msgId,
    istTimestamp,
    String(payload.customerId).trim(),
    String(payload.customerName || "Customer").trim(),
    String(payload.sender || "CUSTOMER").trim(),
    String(payload.messageText || "").trim(),
    false
  ]);

  return { status: "SUCCESS", messageId: msgId };
}

function getChatMessages(customerId) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chat_Messages");
  const data = sheet.getDataRange().getValues();
  const msgs = [];
  const cleanId = String(customerId || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][2]).trim() === cleanId) {
      msgs.push({
        messageId: data[i][0],
        timestamp: data[i][1],
        sender: data[i][4],
        text: data[i][5]
      });
    }
  }
  return { status: "SUCCESS", messages: msgs };
}

function getAllChatThreads(pin) {
  if (!verifyAdminPin(pin)) return { status: "ERROR", message: "Unauthorized PIN." };
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Chat_Messages");
  const data = sheet.getDataRange().getValues();
  const threads = {};

  for (let i = 1; i < data.length; i++) {
    const custId = String(data[i][2]).trim();
    if (!threads[custId]) {
      threads[custId] = {
        customerId: custId,
        customerName: String(data[i][3]),
        lastMessage: String(data[i][5]),
        lastTimestamp: data[i][1],
        sender: data[i][4],
        unreadCount: (data[i][4] === "CUSTOMER" && data[i][6] === false) ? 1 : 0
      };
    } else {
      threads[custId].lastMessage = String(data[i][5]);
      threads[custId].lastTimestamp = data[i][1];
      threads[custId].sender = data[i][4];
      if (data[i][4] === "CUSTOMER" && data[i][6] === false) {
        threads[custId].unreadCount++;
      }
    }
  }

  return { status: "SUCCESS", threads: Object.values(threads).reverse() };
}

function forgotPassword(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();
  const phone = String(payload.phone || "").replace(/\D/g, '').slice(-10);
  const email = String(payload.email || "").trim().toLowerCase();

  for (let i = 1; i < data.length; i++) {
    const rowPhone = String(data[i][2]).replace(/\D/g, '').slice(-10);
    const rowEmail = String(data[i][5]).trim().toLowerCase();

    if (rowPhone === phone && rowEmail === email) {
      const code = Math.floor(100000 + Math.random() * 900000).toString();
      const expires = new Date(Date.now() + 15 * 60 * 1000); // 15 minutes
      sheet.getRange(i + 1, 11).setValue(code);
      sheet.getRange(i + 1, 12).setValue(expires);

      try {
        GmailApp.sendEmail(email, "Yaady's Password Reset Code", `Your password reset code is: ${code}. It expires in 15 minutes.`);
      } catch (e) {}

      return { status: "SUCCESS", message: "Reset code dispatched to your email." };
    }
  }
  return { status: "ERROR", message: "No matching account found with that phone and email." };
}

function resetPassword(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();
  const phone = String(payload.phone || "").replace(/\D/g, '').slice(-10);
  const code = String(payload.resetCode || "").trim();
  const newPass = String(payload.newPassword || "").trim();

  if (newPass.length < 4) return { status: "ERROR", message: "Password must be at least 4 characters." };

  for (let i = 1; i < data.length; i++) {
    const rowPhone = String(data[i][2]).replace(/\D/g, '').slice(-10);
    if (rowPhone === phone) {
      const savedCode = String(data[i][10]).trim();
      const expires = new Date(data[i][11]);

      if (savedCode === code && new Date() <= expires) {
        sheet.getRange(i + 1, 4).setValue(hashString(newPass));
        sheet.getRange(i + 1, 11).setValue("");
        sheet.getRange(i + 1, 12).setValue("");
        return { status: "SUCCESS", message: "Password updated successfully! Please log in." };
      } else {
        return { status: "ERROR", message: "Invalid or expired recovery code." };
      }
    }
  }
  return { status: "ERROR", message: "Account not found." };
}

function changePassword(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Customers");
  const data = sheet.getDataRange().getValues();
  const phone = String(payload.phone || "").replace(/\D/g, '').slice(-10);
  const currHash = hashString(String(payload.currentPassword || "").trim());
  const newPass = String(payload.newPassword || "").trim();

  if (newPass.length < 4) return { status: "ERROR", message: "New password must be at least 4 characters." };

  for (let i = 1; i < data.length; i++) {
    const rowPhone = String(data[i][2]).replace(/\D/g, '').slice(-10);
    if (rowPhone === phone) {
      if (String(data[i][3]) === currHash) {
        sheet.getRange(i + 1, 4).setValue(hashString(newPass));
        return { status: "SUCCESS", message: "Password updated successfully." };
      } else {
        return { status: "ERROR", message: "Current password does not match." };
      }
    }
  }
  return { status: "ERROR", message: "Account not found." };
}