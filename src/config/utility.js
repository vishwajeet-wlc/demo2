/* global Office */

const officeKeys = {
  refreshToken: "microsoft-refresh-token",
  requestForms: "streamline-request-forms",
  requestId: "streamline-request-id",
  clientDomain: "streamline-domain",
  clientToken: "streamline-clientToken",
  selectedFormDetails: "streamline-selectedFormDetails",
};

function setOfficeKeyValue(key, value) {
  Office.context.roamingSettings.set(key, value);
  Office.context.roamingSettings.saveAsync(); // Save changes asynchronously
}

function getOfficeKeyValue(key) {
  return Office.context?.roamingSettings.get(key);
}

function deleteAllKeys() {
  const settings = Office.context.roamingSettings;
  for (var key in settings) {
    settings.remove(key);
  }
  settings.saveAsync();
}

export { setOfficeKeyValue, getOfficeKeyValue, deleteAllKeys, officeKeys };
