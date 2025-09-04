// ==UserScript==
// @name        CopyTime
// @namespace   Violentmonkey Scripts
// @match       https://outlook.office365.com/mail/*
// @match       https://outlook.office.com/mail/*
// @grant       GM_setClipboard
// @grant       GM_setValue
// @grant       GM_getValue
// @version     2.2
// @author      Mental Gravis
// @description 2024. 05. 14. 9:59:28
// @require     https://code.jquery.com/jquery-3.6.4.min.js#sha256-oP6HI9z1XaZNBrJURtCoUT5SUnxFr8s3BzRl+cbzUq8=
// ==/UserScript==
(function() {
    'use strict';

    jQuery(document).ready(function(){

        function copyButtonsSetup() {
            let topMenuRow = document.querySelector("#tablist > div");

            if (topMenuRow) {
                if (!document.querySelector('#CopyQ')) {
                    buttonMaker('CopyQ', topMenuRow, ticketID);
                }
                if (!document.querySelector('#CopyMain')) {
                    buttonMaker('CopyMain', topMenuRow, ticketDisc);
                }
                if (!document.querySelector('#respPerson')) {
                    buttonMaker('respPerson', topMenuRow, respPerson);
                }
            }
            if (document.querySelector('#CopyMain') && document.querySelector('#CopyQ') && document.querySelector('#respPerson')) {
                clearInterval(interWall);
            }
        }

        function findElementByTextContentXPath(textContent) {
            let xpath = "//*[text()='" + textContent + "']";
            let result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
            return result.singleNodeValue;
        }

        async function ticketDisc() {
            try {
                let ticketD = findElementByTextContentXPath("Bejelentés szövege:");
                if (!ticketD) throw new Error("Ticket description element not found.");

                let ticketDiscText = "";
                try{
                    ticketDiscText = ticketD.parentElement.children[1].textContent.trim();
                } catch (error) {
                    ticketDiscText = ticketD.parentElement.parentElement.parentElement.children[1].textContent.trim();
                }
                console.log(ticketDiscText);
                // Call respPerson function to ensure responsible person is handled
                await respPerson();
                GM_setClipboard(ticketDiscText);
            } catch (error) {
                console.error(error);
                alert("Could not find the ticket description.");
            }
        }

        function ticketID() {
            try {
                let ticketI = document.querySelector("#ConversationReadingPaneContainer > div > div > div > div > div > span:nth-child(1)");
                if (!ticketI) throw new Error("Ticket ID element not found.");
                let ticketIDText = ticketI.textContent.match(/Q\d*/g)[0];
                console.log(ticketIDText);
                GM_setClipboard(ticketIDText);
            } catch (error) {
                console.error(error);
                alert("Could not find the ticket ID.");
            }
        }

        function ticketAddr() {
            try {
                let ticketA = findElementByTextContentXPath("Főhelyszín:");
                if (!ticketA) throw new Error("Ticket address element not found.");
                let ticketAddrText = ""
                try{
                    ticketAddrText = ticketA.parentElement.children[3].textContent.trim();
                } catch (error) {
                    ticketAddrText = ticketA.parentElement.parentElement.parentElement.children[3].textContent.trim();
                }
                console.log(ticketAddrText);
                return ticketAddrText;
            } catch (error) {
                console.error(error);
                alert("Could not find the ticket address.");
            }
        }

        async function respPerson() {
            try {
                let address = ticketAddr();
                if (!address) throw new Error("No address found.");

                let responsiblePerson = await GM_getValue(address, null);
                if (responsiblePerson === null) {
                    responsiblePerson = "felelős";
                    await GM_setValue(address, responsiblePerson);
                }
                console.log("Responsible person for " + address + ": " + responsiblePerson);

                // Assuming the title of the email is where ticketID is retrieving it
                let emailTitle = document.querySelector("#ConversationReadingPaneContainer > div > div > div > div > div > span:nth-child(1)");
                if (emailTitle && !emailTitle.textContent.includes(" ------------------> ")) {
                    emailTitle.textContent += " ------------------> " + responsiblePerson;
                }
            } catch (error) {
                console.error(error);
                alert("Could not retrieve or set the responsible person.");
            }
        }

        function buttonMaker(buttonName, buttonPlacement, buttonFunct) {
            let buttonInFunct = document.createElement('button');
            buttonInFunct.textContent = buttonName;
            buttonInFunct.style.cursor = 'pointer';
            buttonInFunct.style.zIndex = '9998';
            buttonInFunct.style.marginLeft = '10px';
            buttonInFunct.style.padding = '6px 12px';
            buttonInFunct.style.border = '1px solid #0078d4';
            buttonInFunct.style.borderRadius = '4px';
            buttonInFunct.style.backgroundColor = '#0078d4';
            buttonInFunct.style.color = '#fff';
            buttonInFunct.style.fontSize = '14px';
            buttonInFunct.style.transition = 'background-color 0.3s ease';

            buttonInFunct.id = buttonName;

            buttonInFunct.addEventListener('mouseover', function() {
                buttonInFunct.style.backgroundColor = "#005a9e";
            });

            buttonInFunct.addEventListener('mouseout', function() {
                buttonInFunct.style.backgroundColor = "#0078d4";
            });

            buttonInFunct.addEventListener('mousedown', function() {
                buttonInFunct.style.backgroundColor = '#004a86';
            });

            buttonInFunct.addEventListener('mouseup', function() {
                buttonInFunct.style.backgroundColor = '#0078d4';
            });

            buttonInFunct.addEventListener('click', buttonFunct);

            buttonPlacement.appendChild(buttonInFunct);
        }

        let interWall = setInterval(copyButtonsSetup, 3000);

    });
})();
