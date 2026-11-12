
/**
 * ============================================================
 * GATEWAY - نظام حركات المخازن
 * ============================================================
 *
 * الوظائف:
 * 1- عرض واجهة Index.
 * 2- تحديد حساب Google للمستخدم.
 * 3- إرسال العمليات إلى Core.
 * 4- تحليل استجابة Core.
 * 5- إظهار رسائل أخطاء واضحة.
 *
 * ============================================================
 */


/* ============================================================
   1. إعدادات الاتصال
   ============================================================ */

/**
 * ضع هنا رابط نشر Core الحقيقي.
 * يجب أن يكون رابط Web App وينتهي بـ /exec
 */
const CORE_URL =
  'https://script.google.com/macros/s/AKfycbyOEeoqbfm4uHYJLOi5AZ7Dm_bhU7sl7m1RuZbtLVBvMDCeodWhXFZF4lTKW5rWwiNC/exec';


/* ============================================================
   2. عرض الواجهة الرئيسية
   ============================================================ */

function doGet() {

  return HtmlService
    .createHtmlOutputFromFile('Index')
    .setTitle('حركات المخازن')
    .addMetaTag(
      'viewport',
      'width=device-width, initial-scale=1'
    );

}


/* ============================================================
   3. قراءة المفتاح السري
   ============================================================ */

function getGatewaySecret_() {

  const secret = PropertiesService
    .getScriptProperties()
    .getProperty('GATEWAY_SECRET');

  if (!secret) {
    throw new Error(
      'مفتاح الاتصال GATEWAY_SECRET غير موجود.'
    );
  }

  return secret;

}


/* ============================================================
   4. استخراج بريد Google
   ============================================================ */

function getCurrentGoogleEmail_() {

  const email = Session
    .getActiveUser()
    .getEmail();

  if (!email) {

    throw new Error(
      'تعذر التعرف على حساب Google. ' +
      'تأكد من إعداد نشر التطبيق.'
    );

  }

  return String(email)
    .toLowerCase()
    .trim();

}


/* ============================================================
   5. التحقق من رابط Core
   ============================================================ */

function validateCoreUrl_() {

  if (
    !CORE_URL ||
    CORE_URL.indexOf('https://script.google.com/macros/s/') !== 0 ||
    !/\/exec$/.test(CORE_URL) ||
    CORE_URL.indexOf('ضع_رابط') !== -1
  ) {

    throw new Error(
      'رابط Core غير صحيح. ' +
      'ضع رابط نشر Core الكامل الذي ينتهي بـ /exec.'
    );

  }

}


/* ============================================================
   6. إرسال الطلب إلى Core
   ============================================================ */

function sendToCore_(email, action, payload) {

  validateCoreUrl_();

  const secret = getGatewaySecret_();

  const requestBody = {

    secret: secret,

    email: email,

    action: action,

    payload: payload || {}

  };


  let response;

  try {

    response = UrlFetchApp.fetch(CORE_URL, {

      method: 'post',

      contentType: 'application/json',

      payload: JSON.stringify(requestBody),

      muteHttpExceptions: true,

      followRedirects: true

    });

  } catch (error) {

    throw new Error(
      'تعذر الاتصال بعنوان Core: ' +
      error.message
    );

  }


  const status = response.getResponseCode();

  const body = response.getContentText();


  // ----------------------------------------------------------
  // HTTP 401
  // ----------------------------------------------------------

  if (status === 401) {

    throw new Error(
      'رفض Google الوصول إلى Core (HTTP 401). ' +
      'راجع إعدادات نشر Core: ' +
      'Execute as = Me، وتأكد من أن الوصول مسموح للبوابة. ' +
      'بداية الاستجابة: ' +
      body.substring(0, 250)
    );

  }


  // ----------------------------------------------------------
  // HTTP 403
  // ----------------------------------------------------------

  if (status === 403) {

    throw new Error(
      'الوصول إلى Core ممنوع (HTTP 403). ' +
      'راجع صلاحيات النشر وسياسات حساب Google.'
    );

  }


  // ----------------------------------------------------------
  // HTTP 404
  // ----------------------------------------------------------

  if (status === 404) {

    throw new Error(
      'لم يتم العثور على Core (HTTP 404). ' +
      'راجع رابط النشر.'
    );

  }


  // ----------------------------------------------------------
  // أخطاء HTTP الأخرى
  // ----------------------------------------------------------

  if (status !== 200) {

    throw new Error(
      'فشل الاتصال بخادم Core. ' +
      'HTTP: ' + status +
      ' | الاستجابة: ' +
      body.substring(0, 500)
    );

  }


  // ----------------------------------------------------------
  // تحليل JSON
  // ----------------------------------------------------------

  let result;

  try {

    result = JSON.parse(body);

  } catch (error) {

    throw new Error(
      'أعاد Core استجابة ليست JSON. ' +
      'قد تكون صفحة تسجيل دخول أو صلاحيات Google. ' +
      'بداية الاستجابة: ' +
      body.substring(0, 500)
    );

  }


  // ----------------------------------------------------------
  // التحقق من شكل الاستجابة
  // ----------------------------------------------------------

  if (
    !result ||
    typeof result !== 'object' ||
    Array.isArray(result)
  ) {

    throw new Error(
      'صيغة استجابة Core غير صحيحة.'
    );

  }


  if (result.ok !== true) {

    throw new Error(
      result.error ||
      'رفض Core تنفيذ العملية.'
    );

  }


  return result.data;

}


/* ============================================================
   7. واجهة API الرئيسية
   ============================================================ */

/**
 * تستدعيها واجهة Index الحالية.
 *
 * call('meta')
 * call('list')
 * call('available')
 * call('submit', payload)
 * call('edit', payload)
 * call('cancel', payload)
 * call('approve', payload)
 * call('reject', payload)
 */

function api(action, payloadJson) {

  try {

    // التحقق من اسم العملية

    if (
      !action ||
      typeof action !== 'string'
    ) {

      throw new Error(
        'اسم العملية غير صالح.'
      );

    }


    // تحليل البيانات

    let payload = {};

    try {

      payload = JSON.parse(
        payloadJson || '{}'
      );

    } catch (error) {

      throw new Error(
        'بيانات العملية غير صالحة.'
      );

    }


    // استخراج بريد المستخدم

    const email = getCurrentGoogleEmail_();


    // إرسال الطلب إلى Core

    return sendToCore_(
      email,
      action,
      payload
    );


  } catch (error) {

    console.error(
      'Gateway API Error: ' + error.message
    );

    throw new Error(
      error.message
    );

  }

}


/* ============================================================
   8. اختبار إعدادات Gateway
   ============================================================ */

function gatewaySelfTest() {

  const result = {

    gateway: 'ONLINE',

    timestamp: new Date().toISOString(),

    coreUrlConfigured: false,

    secretConfigured: false,

    googleEmailDetected: false,

    email: null

  };


  try {

    validateCoreUrl_();

    result.coreUrlConfigured = true;

  } catch (error) {

    result.coreUrlError = error.message;

  }


  try {

    const secret = getGatewaySecret_();

    result.secretConfigured = Boolean(secret);

  } catch (error) {

    result.secretError = error.message;

  }


  try {

    const email = getCurrentGoogleEmail_();

    result.googleEmailDetected = true;

    result.email = email;

  } catch (error) {

    result.googleEmailError = error.message;

  }


  Logger.log(
    JSON.stringify(result, null, 2)
  );

  return result;

}


/* ============================================================
   9. اختبار اتصال Core
   ============================================================ */

/**
 * اختبار الاتصال باستخدام عملية meta.
 *
 * لا يغير بيانات المخازن.
 */

function testCoreConnection() {

  const email = getCurrentGoogleEmail_();

  const result = sendToCore_(
    email,
    'meta',
    {}
  );

  Logger.log(
    'تم الاتصال بـ Core بنجاح.'
  );

  return result;

}
