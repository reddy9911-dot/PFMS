/**
 * PFMS - "Entered By" names.
 *
 * Set the names shown on the payment forms without editing code:
 *   Project Settings > Script Properties > Add property
 *     Property: PFMS_PEOPLE
 *     Value:    comma-separated names, e.g.  Alex, Maria
 *
 * Until it is set, the forms show the placeholder names in App.html
 * (PFMS_DEFAULT_PEOPLE).
 *
 * The chosen name is sent as payment.enteredBy to saveReceivePayment and
 * saveMakePayment (Code.gs), which save it in the payment row.
 */

/**
 * @return {string[]} names to show in the Entered By picker, or [] if unset.
 */
function getPfmsPeople() {
  var value = PropertiesService.getScriptProperties().getProperty('PFMS_PEOPLE') || '';
  return value
    .split(',')
    .map(function (name) {
      return name.trim();
    })
    .filter(function (name) {
      return name.length > 0;
    });
}
