(function (window, undefined) {
    'use strict';

    function addTagToSelection(tagValue, commentText) {
        var controlId = Date.now();

        window.Asc.plugin.executeMethod('AddContentControl', [1, {
            Id: controlId,
            Lock: 0,
            Tag: tagValue,
            Alias: tagValue,
            Appearance: 1
        }], function () {
            window.Asc.plugin.executeMethod('AddComment', [{
                Text: commentText,
                UserName: 'Policy System',
                UserId: 'policy-system',
                Time: Date.now(),
                Solved: false
            }]);
        });
    }

    function buildMenuItems(options) {
        if (!options) {
            return null;
        }

        if (options.type !== 'Selection' && options.type !== 'Target') {
            return null;
        }

        return {
            guid: window.Asc.plugin.guid,
            items: [
                {
                    id: 'policy_tagging_root',
                    text: 'Policy Tagging',
                    items: [
                        {
                            id: 'policy_tag_review',
                            text: 'Tag Paragraph: Needs Review'
                        },
                        {
                            id: 'policy_tag_compliant',
                            text: 'Tag Paragraph: Compliant'
                        }
                    ]
                }
            ]
        };
    }

    window.Asc.plugin.init = function () {
        // No visual UI. This plugin only contributes right-click actions.
    };

    window.Asc.plugin.button = function () {};

    window.Asc.plugin.attachEvent('onContextMenuShow', function (options) {
        var menuItems = buildMenuItems(options);
        if (!menuItems) {
            return;
        }

        window.Asc.plugin.executeMethod('AddContextMenuItem', [menuItems]);
    });

    window.Asc.plugin.attachContextMenuClickEvent('policy_tag_review', function () {
        addTagToSelection('{policy:needs-review}', '[POLICY TAG] Paragraph tagged as Needs Review');
    });

    window.Asc.plugin.attachContextMenuClickEvent('policy_tag_compliant', function () {
        addTagToSelection('{policy:compliant}', '[POLICY TAG] Paragraph tagged as Compliant');
    });

})(window, undefined);
