/**
 * (c) Copyright Ascensio System SIA 2020
 * Licensed under the Apache License, Version 2.0
 * http://www.apache.org/licenses/LICENSE-2.0
 */
(function(window, undefined){
    var flagInit = false;
    var fBtnGetAll = false;
    var ArrContentControls = {};

    window.Asc.plugin.init = function() {
        document.getElementById("buttonIDGetAll").onclick = function() {
            window.Asc.plugin.executeMethod("GetAllContentControls");
            fBtnGetAll = true;
        };

        if (!flagInit) {
            flagInit = true;
            window.Asc.plugin.executeMethod("GetAllContentControls");
        }
    };

    addLabel = (arrEl, element) => {
        var fClickLabel = false;
        $(element).append(
            $('<label>', {
                id: arrEl.tag,
                for: element,
                class: 'label-info',
                text: arrEl.tag,
                on: {
                    click: function() {
                        fClickLabel = true;
                        $('.label-selected').removeClass('label-selected');
                        $(this).addClass('label-selected');
                        // Jump to and select the first content control matching this tag,
                        // so clicking a label actually navigates to the tagged paragraph.
                        window.Asc.plugin.executeMethod("SelectContentControl", [arrEl.id[0]]);
                    },
                    mouseover: function() {
                        $(this).addClass('label-hovered');
                    },
                    mouseout: function() {
                        $(this).removeClass('label-hovered');
                    }
                }
            })
        );
    };

    compareArr = (arr) => {
        ArrContentControls = {};
        for (var i = 0; i < arr.length; i++) {
            if (!arr[i].Tag) {
                continue;
            }
            if (ArrContentControls[arr[i].Tag]) {
                ArrContentControls[arr[i].Tag].id.push(arr[i].InternalId);
            } else {
                ArrContentControls[arr[i].Tag] = {
                    id: [arr[i].InternalId],
                    tag: arr[i].Tag
                };
            }
        }
    };

    window.Asc.plugin.button = function() {
        this.executeCommand("close", "");
    };

    window.Asc.plugin.onMethodReturn = function(returnValue) {
        var _plugin = window.Asc.plugin;
        if (_plugin.info.methodName == "GetAllContentControls") {
            compareArr(returnValue);
            fBtnGetAll = false;
            document.getElementById("divG").innerHTML = "";
            for (const key in ArrContentControls) {
                addLabel(ArrContentControls[key], "#divG");
            }
        }
    };
})(window, undefined);
