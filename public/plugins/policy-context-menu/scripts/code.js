(function (window, undefined) {
    'use strict';

    // Highlight colors so tagged paragraphs are visually distinguishable at a glance.
    var TAG_COLORS = {
        'needs-review': [255, 235, 156],
        'compliant': [198, 239, 206],
        'custom': [198, 224, 255]
    };

    // Wraps the whole paragraph under the cursor in a block-level content control (ApiBlockLvlSdt)
    // via the Document Builder API, so the tag survives as part of the .docx itself.
    function addTagToSelection(tagValue, commentText, colorKey) {
        window.Asc.scope.tagValue = tagValue;
        window.Asc.scope.commentText = commentText;
        window.Asc.scope.color = TAG_COLORS[colorKey] || TAG_COLORS.custom;

        window.Asc.plugin.callCommand(function () {
            try {
                var oDocument = Api.GetDocument();
                var oRange = oDocument.GetRangeBySelect();
                var oParagraph = oRange ? oRange.GetParagraph(0) : null;

                if (!oParagraph) {
                    console.warn('[Policy Tagging] No paragraph found — place the cursor inside a paragraph and try again.');
                    return;
                }

                // Comment is best-effort audit trail; its failure must not block the visible tag/highlight.
                try {
                    oRange.AddComment(Asc.scope.commentText, 'Policy System');
                } catch (commentErr) {
                    console.warn('[Policy Tagging] Could not add audit comment:', commentErr);
                }

                // Shade the paragraph background so the tag is visible without opening the tags panel.
                var c = Asc.scope.color;
                oParagraph.SetShd('clear', Api.RGB(c[0], c[1], c[2]));

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
                        }
                    ]
                }
            ]
        };
    }

    window.Asc.plugin.init = function () {
        var input = document.getElementById('customTagInput');
        var btn = document.getElementById('customTagBtn');

        btn.onclick = function () {
            var tagName = (input.value || '').trim();
            if (!tagName) {
                input.focus();
                return;
            }

            addTagToSelection('{policy:' + tagName + '}', '[POLICY TAG] Paragraph tagged as "' + tagName + '"', 'custom');
            input.value = '';
        };
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
        addTagToSelection('{policy:needs-review}', '[POLICY TAG] Paragraph tagged as Needs Review', 'needs-review');
    });

    window.Asc.plugin.attachContextMenuClickEvent('policy_tag_compliant', function () {
        addTagToSelection('{policy:compliant}', '[POLICY TAG] Paragraph tagged as Compliant', 'compliant');
    });

})(window, undefined);
