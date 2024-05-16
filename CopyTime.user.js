// ==UserScript==
// @name        CopyTime
// @namespace   Violentmonkey Scripts
// @match       https://outlook.office.com/mail/*
// @grant       GM_setClipboard
// @version     1.0
// @author      Mental Gravis
// @description 2024. 05. 14. 9:59:28
// @require     https://code.jquery.com/jquery-3.6.4.min.js#sha256-oP6HI9z1XaZNBrJURtCoUT5SUnxFr8s3BzRl+cbzUq8=
// ==/UserScript==
(function() {
    'use strict';
    // console.log("outerrun");
    jQuery(document).ready(function(){
        // console.log('fut');

        function copyButtonsSetup () {
            //console.log('CopyButtonSetup');

            let topMenuRow = document.querySelector("#tablist > div");

            if (topMenuRow){
                // console.log('topMenuRow is True');
                if (!(document.querySelector('#CopyQ'))){
                    buttonMaker('CopyQ', topMenuRow, ticketID);
                }
                if (!(document.querySelector('#CopyMain'))){
                    buttonMaker('CopyMain', topMenuRow, ticketDisc);
                }
            }
            if ((document.querySelector('#CopyMain')) && (document.querySelector('#CopyQ'))){
                clearInterval(interWall);
            }

        }

        function findElementByTextContentXPath(textContent) {
            let xpath = "//*[text()='" + textContent + "']";
            let result = document.evaluate(xpath, document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null);
            return result.singleNodeValue;
        }

        function ticketDisc () {
            let ticketD = findElementByTextContentXPath("Bejelentés szövege:");
            let ticketDiscText = ticketD.parentElement.children[1].textContent.trim();
            console.log(ticketDiscText);
            GM_setClipboard(ticketDiscText);
        }

        function ticketID () { // gets the Q number from title of the message
            let ticketI = document.querySelector("#ConversationReadingPaneContainer > div > div > div > div > div > div > div > div > div > span:nth-child(1)");
            let ticketIDText = ticketI.textContent.match(/Q\d*/g)[0];
            console.log(ticketIDText);
            GM_setClipboard(ticketIDText);
        }

        function buttonMaker (buttonName, buttonPlacement, buttonFunct) {
            let buttonInFunct = document.createElement('button');
            buttonInFunct.textContent = buttonName;
            // buttonInFunct.style.display = 'block';
            buttonInFunct.style.cursor = 'pointer';
            buttonInFunct.style.zIndex = '9998'; // Set a high z-index value for the button
            // buttonInFunct.style.MarginLeft = '10px';
            buttonInFunct.style.paddingRight = '12px';
            buttonInFunct.style.paddingLeft = '12px';
            buttonInFunct.style.background = 'transparent';
            buttonInFunct.style.transition = 'background-color 0s ease-in-out';
            buttonInFunct.id = buttonName;


            buttonInFunct.addEventListener('mouseover', function(){
                buttonInFunct.style.backgroundColor = "#f0f0f0";
                buttonInFunct.style.textDecoration = "underline";
            });

            buttonInFunct.addEventListener('mouseout', function() {
                buttonInFunct.style.backgroundColor = "transparent";
                buttonInFunct.style.textDecoration = "none";
            });

            buttonInFunct.addEventListener('mousedown', function() {
                buttonInFunct.style.backgroundColor = '#c6e1e2';
            });

            buttonInFunct.addEventListener('mouseup', function() {
                buttonInFunct.style.backgroundColor = 'transparent';
            });

            buttonInFunct.addEventListener('click', buttonFunct);

            buttonPlacement.appendChild(buttonInFunct);
        }

        let interWall = setInterval(copyButtonsSetup, 3000);

    });
})();
