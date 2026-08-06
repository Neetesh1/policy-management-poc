(function (window, undefined) {
    'use strict';

    // Highlight colors so tagged paragraphs are visually distinguishable at a glance.
    var TAG_COLORS = {
        'needs-review': [255, 235, 156],
        'compliant': [198, 239, 206],
        'custom': [198, 224, 255],
        'regulation': [225, 213, 245]
    };

    // Wraps the whole paragraph under the cursor in a block-level content control (ApiBlockLvlSdt)
    // via the Document Builder API, so the tag survives as part of the .docx itself.
    function addTagToSelection(tagValue, colorKey, aliasLabel) {
        window.Asc.scope.tagValue = tagValue;
        window.Asc.scope.color = TAG_COLORS[colorKey] || TAG_COLORS.custom;
        window.Asc.scope.aliasLabel = aliasLabel;

        window.Asc.plugin.callCommand(function () {
            try {
                var oDocument = Api.GetDocument();
                var oRange = oDocument.GetRangeBySelect();
                var oParagraph = oRange ? oRange.GetParagraph(0) : null;

                if (!oParagraph) {
                    console.warn('[Policy Tagging] No paragraph found — place the cursor inside a paragraph and try again.');
                    return;
                }

                // Audit trail (who/when/label) is persisted server-side via onChangeContentControl
                // in editor.html rather than as an in-document Word comment.

                // Shade the paragraph background so the tag is visible without opening the tags panel.
                var c = Asc.scope.color;
                oParagraph.SetShd('clear', Api.RGB(c[0], c[1], c[2]));

                // Paragraph is already part of the document tree, so it must be wrapped in place
                // rather than pushed into a freshly created content control (Push/InsertContent
                // only works for detached elements not yet added to the document).
                var blockLvlSdt = oParagraph.InsertInContentControl(1);
                blockLvlSdt.SetTag(Asc.scope.tagValue);
                // Alias is the content control's native "Title": Word/ONLYOFFICE show it as a
                // floating tab above the control when the cursor is placed inside it.
                blockLvlSdt.SetAlias(Asc.scope.aliasLabel);
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
        var reviewBtn = document.getElementById('quickTagReview');
        var compliantBtn = document.getElementById('quickTagCompliant');
        var regulationSelect = document.getElementById('regulationSelect');
        var assignRegulationBtn = document.getElementById('assignRegulationBtn');

        // Populate the regulation dropdown from the app's catalog (same origin as this plugin).
        fetch('/api/regulations')
            .then(function (res) { return res.json(); })
            .then(function (regulations) {
                regulationSelect.innerHTML = '';
                regulations.forEach(function (r) {
                    var opt = document.createElement('option');
                    opt.value = r.code;
                    opt.textContent = r.code + ' — ' + r.name;
                    opt.dataset.name = r.name;
                    regulationSelect.appendChild(opt);
                });
            })
            .catch(function (err) {
                console.warn('[Policy Tagging] Failed to load regulations:', err);
                regulationSelect.innerHTML = '<option value="">Failed to load regulations</option>';
            });

        assignRegulationBtn.onclick = function () {
            var code = regulationSelect.value;
            if (!code) {
                regulationSelect.focus();
                return;
            }
            var opt = regulationSelect.options[regulationSelect.selectedIndex];
            var name = (opt && opt.dataset.name) || code;
            addTagToSelection('{regulation:' + code + '}', 'regulation', code + ' — ' + name);
        };

        reviewBtn.onclick = function () {
            addTagToSelection('{policy:needs-review}', 'needs-review', 'Needs Review');
        };

        compliantBtn.onclick = function () {
            addTagToSelection('{policy:compliant}', 'compliant', 'Compliant');
        };

        btn.onclick = function () {
            var tagName = (input.value || '').trim();
            if (!tagName) {
                input.focus();
                return;
            }

            addTagToSelection('{policy:' + tagName + '}', 'custom', tagName);
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
        addTagToSelection('{policy:needs-review}', 'needs-review', 'Needs Review');
    });

    window.Asc.plugin.attachContextMenuClickEvent('policy_tag_compliant', function () {
        addTagToSelection('{policy:compliant}', 'compliant', 'Compliant');
    });

})(window, undefined);
