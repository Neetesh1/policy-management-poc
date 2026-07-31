(function (window, undefined) {
    'use strict';

    // Wraps the whole paragraph under the cursor in a block-level content control (ApiBlockLvlSdt)
    // via the Document Builder API, so the tag survives as part of the .docx itself.
    function addTagToSelection(tagValue, commentText) {
        window.Asc.scope.tagValue = tagValue;
        window.Asc.scope.commentText = commentText;

        window.Asc.plugin.callCommand(function () {
            try {
                var oDocument = Api.GetDocument();
                var oRange = oDocument.GetRangeBySelect();
                var oParagraph = oRange ? oRange.GetParagraph(0) : null;

                if (!oParagraph) {
                    console.warn('[Policy Tagging] No paragraph found — place the cursor inside a paragraph and try again.');
                    return;
                }

                oRange.AddComment(Asc.scope.commentText, 'Policy System');

                // Paragraph is already part of the document tree, so it must be wrapped in place
                // rather than pushed into a freshly created content control (Push/InsertContent
                // only works for detached elements not yet added to the document).
                var blockLvlSdt = oParagraph.InsertInContentControl(1);
                blockLvlSdt.SetTag(Asc.scope.tagValue);
            } catch (e) {
                console.error('[Policy Tagging] Failed to tag paragraph:', e);
            }
        }, false);
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
                        },
                        {
                            id: 'policy_tag_custom',
                            text: 'Tag Paragraph: Custom...'
                        }
                    ]
                }
            ]
        };
    }

    function addCustomTagToSelection() {
        var tagName = window.prompt('Enter a custom tag name for the selected paragraph:');
        if (!tagName) {
            return;
        }
        tagName = tagName.trim();
        if (!tagName) {
            return;
        }

        addTagToSelection('{policy:' + tagName + '}', '[POLICY TAG] Paragraph tagged as "' + tagName + '"');
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

    window.Asc.plugin.attachContextMenuClickEvent('policy_tag_custom', function () {
        addCustomTagToSelection();
    });

})(window, undefined);
