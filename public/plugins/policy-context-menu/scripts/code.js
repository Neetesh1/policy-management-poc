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
                    return null;
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
                return blockLvlSdt.GetId();
            } catch (e) {
                console.error('[Policy Tagging] Failed to tag paragraph:', e);
                return null;
            }
        }, false, false, function (contentControlId) {
            if (contentControlId) {
                attachTagButton(contentControlId);
            }
        });
    }

    // Custom button on the content control itself (AddContentControlButtons, since ONLYOFFICE 9.0)
    // instead of relying on the native gray Alias tab — clicking it opens our own "Policy Tag" card
    // below rather than a plain tooltip.
    function attachTagButton(contentControlId) {
        var buttons = { guid: window.Asc.plugin.guid, items: {} };
        buttons.items[contentControlId] = [{
            id: 'policyTagCard',
            // Reuses this plugin's own icon assets (light/dark, per-DPI) registered in config.json.
            icons: 'resources/%theme-type%(light|dark)/icon%scale%(100|125|150|175|200).png'
        }];
        window.Asc.plugin.executeMethod('AddContentControlButtons', [buttons]);
    }

    function onContentControlButtonClick(data) {
        if (!data || data.buttonId !== 'policyTagCard' || !data.contentControlId) return;

        window.Asc.scope.ccId = data.contentControlId;
        window.Asc.plugin.callCommand(function () {
            try {
                var ccs = Api.GetDocument().GetAllContentControls();
                for (var i = 0; i < ccs.length; i++) {
                    if (ccs[i].GetId() === Asc.scope.ccId) {
                        return { tag: ccs[i].GetTag(), alias: ccs[i].GetAlias() };
                    }
                }
                return null;
            } catch (e) {
                return null;
            }
        }, false, false, function (info) {
            showTagCard(data.contentControlId, info);
        });
    }

    function parseTagType(tag) {
        if (/^\{regulation:/.test(tag || '')) return 'Regulation';
        if (/^\{policy:needs-review\}$/.test(tag || '')) return 'Needs Review';
        if (/^\{policy:compliant\}$/.test(tag || '')) return 'Compliant';
        if (/^\{policy:/.test(tag || '')) return 'Custom Tag';
        return 'Unknown';
    }

    function showTagCard(contentControlId, info) {
        var card = document.getElementById('tagCard');
        var fields = document.getElementById('tagCardFields');

        if (!info) {
            fields.innerHTML = '<dd style="color:#c0392b;">Could not read tag details.</dd>';
        } else {
            fields.innerHTML =
                '<dt>Status</dt><dd>' + (info.alias || '—') + '</dd>' +
                '<dt>Type</dt><dd>' + parseTagType(info.tag) + '</dd>' +
                '<dt>Tag Value</dt><dd style="font-family:monospace;font-size:10px;">' + (info.tag || '—') + '</dd>';
        }

        card.style.display = 'block';
        document.getElementById('tagCardGoTo').onclick = function () {
            window.Asc.plugin.executeMethod('SelectContentControl', [contentControlId]);
        };
        document.getElementById('tagCardRemove').onclick = function () {
            removeTag(contentControlId);
            card.style.display = 'none';
        };
    }

    // Clears the paragraph shading first (RemoveContentControl only unwraps the tag, it
    // doesn't know about the background color we set separately), then unwraps the control.
    function removeTag(contentControlId) {
        window.Asc.scope.ccId = contentControlId;
        window.Asc.plugin.callCommand(function () {
            try {
                var ccs = Api.GetDocument().GetAllContentControls();
                for (var i = 0; i < ccs.length; i++) {
                    if (ccs[i].GetId() === Asc.scope.ccId) {
                        var oParagraph = ccs[i].GetContent().GetElement(0);
                        if (oParagraph && oParagraph.SetShd) oParagraph.SetShd('nil');
                        break;
                    }
                }
            } catch (e) {
                console.error('[Policy Tagging] Failed to clear shading before removing tag:', e);
            }
        }, false, false, function () {
            window.Asc.plugin.executeMethod('RemoveContentControl', [contentControlId]);
        });
    }


    // Extracts the structured added/removed change list from the last Compare Versions
    // run (ApiDocument.GetReviewReport), as an alternative to the visual redline only.
    function renderReviewReport(report) {
        var list = document.getElementById('reviewReportList');
        var changes = Array.isArray(report) ? report : (report && Array.isArray(report.Changes) ? report.Changes : null);

        if (!changes) {
            list.innerHTML = '<p style="color:#94a3b8;text-align:center;padding:8px 0;">No report data returned — run Compare Versions first.</p>';
            return;
        }
        if (changes.length === 0) {
            list.innerHTML = '<p style="color:#94a3b8;text-align:center;padding:8px 0;">No changes found.</p>';
            return;
        }

        list.innerHTML = changes.map(function (c) {
            var type = c.Type || c.type || c.ReviewType || 'change';
            var text = c.Text || c.text || c.Value || JSON.stringify(c);
            var user = c.User || c.user || c.Author || '';
            var isRemoved = /remov|delet/i.test(String(type));
            var color = isRemoved ? '#c0392b' : '#2e7d32';
            var prefix = isRemoved ? '−' : '+';
            return '<div style="padding:4px 0;border-bottom:1px solid #f0f2f6;color:' + color + ';">' +
                '<strong>' + prefix + '</strong> ' + String(text).replace(/</g, '&lt;') +
                (user ? ' <span style="color:#94a3b8;">(' + String(user).replace(/</g, '&lt;') + ')</span>' : '') +
                '</div>';
        }).join('');
    }

    function showReviewReport() {
        var list = document.getElementById('reviewReportList');
        list.innerHTML = '<p style="color:#94a3b8;text-align:center;padding:8px 0;">Loading…</p>';

        window.Asc.plugin.callCommand(function () {
            try {
                return Api.GetDocument().GetReviewReport();
            } catch (e) {
                return { error: String(e) };
            }
        }, false, false, function (result) {
            if (result && result.error) {
                list.innerHTML = '<p style="color:#c0392b;text-align:center;padding:8px 0;">GetReviewReport failed: ' + result.error.replace(/</g, '&lt;') + '</p>';
                return;
            }
            renderReviewReport(result);
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
        window.Asc.plugin.attachEvent('onContentControlButtonClick', onContentControlButtonClick);

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

        document.getElementById('showReviewReportBtn').onclick = showReviewReport;

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
