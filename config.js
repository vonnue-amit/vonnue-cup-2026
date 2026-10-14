/*
 * Site settings. This is the only file you need to edit to change where results come from.
 *
 * To read results straight from Google Sheets, paste the sheet's normal link (the one in your
 * browser's address bar) between the quotes below. In Google Sheets, set
 * Share → General access → "Anyone with the link" → Viewer, or the site can't read it.
 * Leave it empty to use data/Vonnue_Cup_Knockout_Brackets.xlsx instead.
 */
window.VONNUE_CONFIG = {
  // Results workbook (the brackets sheet), e.g. 'https://docs.google.com/spreadsheets/d/1AbC.../edit'
  googleSheet: 'https://docs.google.com/spreadsheets/d/16AgRzH3_pqWp6hOs5vzFkgxl9DDVl5D7pRjhFYUBlO8/edit',

  // Optional: the participants workbook, if you also keep it in Google Sheets
  googleParticipantsSheet: '',

  // The workbook dates have no year
  year: 2026,

  // How often open pages check for new results, in minutes
  refreshMinutes: 2,
};
