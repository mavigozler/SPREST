"use strict";

import { SPServerREST } from "../src/SPRESTSupportLib";
import { SPSiteREST } from "../src/SPSiteREST";
import { SPListREST } from "../src/SPListREST";

const SPServer = "http://mhalloran.sharepoint.com",
   SPSite = "/sites/mhteam/Projects";

function maintest() {
/* Test plan */
// 0. get server
   const spServer = new SPServerREST({URL: SPServer});
   spServer.getSiteProperties().then((response) => {
      console.log(JSON.stringify(response, null, "  "))
   }).catch((err) => {
      console.log(`ERROR: SPSiteREST.getSiteProperties()\n  ${err.message}`);
   });
   spServer.getWebProperties().then((response) => {
      console.log(JSON.stringify(response, null, "  "))
   }).catch((err) => {
      console.log(`ERROR: SPSiteREST.getWebProperties()\n  ${err.message}`);
   });
   spServer.getRootweb().then((response) => {
      console.log(JSON.stringify(response, null, "  "))
   }).catch((err) => {
      console.log(`ERROR: SPSiteREST.getRootweb()\n  ${err.message}`);
   });
// 1. Create a site

   const spSite = new SPSiteREST({
      server: SPServer,
      site: SPSite
   });
   spSite.createSubsite(
      "Test Subsite",
      "Just a test",
      "testsubsite",
      "STS#0",
      false
   ).then((response) => {
      console.log("OK " + response);
   }).catch((err) => {
      console.log("ERROR: " + err);
   });
   // verify the creation ...how?

// 2. Create a list
   spSite.createList(
      "Test list",
      100,
      "Testing list creation",
      true,
      true
   ).then((response) => {

   }).catch((err) => {

   });

   const spList = new SPListREST({
      server: SPServer,
      site: SPSite,
      listGuid: ""
   });

   //spList.createFields({});

// 3. Create columns in list
// 4. Create two items
// 5. Create a library
// 6. Create library columns
// 7. Create create two library items
// 8. Check out one library item
// 9. Change metadata on library item
// 10.  Check in that item
// 11. Delete one list item
// 12. Create another list
// 13. Update the list
// 14. Delete the list
// 15. Create a site
// 16. Update the site
// 17. Delete the site

}

maintest();