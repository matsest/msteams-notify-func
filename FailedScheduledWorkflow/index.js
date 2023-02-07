const axios = require('axios'); // axios must be installed via npm i axios TODO: handle in func deps
const Crypto = require('crypto');

// TODO - get this as variable
const webhookURL = ""

module.exports = async function (context, req) {
    context.log('JavaScript HTTP trigger function processed a request.');

    const hmac = Crypto.createHmac("sha1", ""); // TODO - get from secret
    const signature = hmac.update(JSON.stringify(req.body)).digest('hex');
    const shaSignature = `sha1=${signature}`;
    const gitHubSignature = req.headers['x-hub-signature'];

    const name = (req.query.name || (req.body && req.body.name));

    if (!shaSignature.localeCompare(gitHubSignature)) {
        // https://docs.github.com/en/rest/actions/workflow-runs?apiVersion=2022-11-28#get-a-workflow-run
        if (req.body.action) {
            var workflowStatus = req.body.workflow_run.conclusion;
            if (workflowStatus == "failure") {
                context.log(postMessageToTeams(
                    "Scheduled workflow failed!",
                    req.body.workflow_run.name,
                    req.body.workflow_run.updated_at,
                    req.body.workflow_run.html_url
                )
                );
                context.res = {
                    status: 200,
                    body: {
                        name: req.body.workflow_run.name,
                        event: req.body.event,
                        state: req.body.action,
                        timestamp: req.body.updated_at,
                        conclusion: req.body.workflow_run.conclusion,
                        type: req.headers['x-github-event'],
                        url: req.body.workflow_run.html_url
                    }
                };
            }
            else {
                context.res = {
                    status: 200,
                    body: ("Workflow succeeded")
                };
            };
        }
        else {
            context.res = {
                status: 400,
                body: ("Invalid payload for actions event")
            };
        }
    }
    else {
        context.res = {
            status: 401,
            body: "Signatures don't match"
        };
    }
}

// TODO: simplify load from json?
async function postMessageToTeams(title, message, time, url) {
    const card = {
        "type": "message",
        "attachments": [
            {
                "contentType": "application/vnd.microsoft.card.adaptive",
                "contentUrl": null,
                "content": {
                    "$schema": "http://adaptivecards.io/schemas/adaptive-card.json",
                    "type": "AdaptiveCard",
                    "version": "1.2",
                    "body": [
                        {
                            "type": "TextBlock",
                            "size": "Medium",
                            "weight": "Bolder",
                            "text": title,
                            "color": "attention"
                        },
                        {
                            "type": "ColumnSet",
                            "columns": [
                                {
                                    "type": "Column",
                                    "items": [
                                        {
                                            "type": "Image",
                                            "style": "Person",
                                            "url": "", // TODO
                                            "size": "Small"
                                        }
                                    ],
                                    "width": "auto"
                                },
                                {
                                    "type": "Column",
                                    "items": [
                                        {
                                            "type": "TextBlock",
                                            "weight": "Bolder",
                                            "text": "github-actions-bot",
                                            "wrap": true
                                        },
                                        {
                                            "type": "TextBlock",
                                            "spacing": "None",
                                            "text": "Created {{DATE(" + time + ",LONG)}}",
                                            "isSubtle": true,
                                            "wrap": true
                                        }
                                    ],
                                    "width": "stretch"
                                }
                            ]
                        },
                        {
                            "type": "TextBlock",
                            "size": "Medium",
                            "text": "- Workflow name: " + message + "\r- Repo: [demo](https://github.com)\r- Org: [demoorg](https://github.com)" // TODO PARAMS
                        },
                        {
                            "type": "TextBlock",
                            "size": "Medium",
                            "text": "**Please click the link below to investigate, and add comments in this thread to follow up!**",
                            "wrap": true
                        },
                    ],
                    "actions": [
                        {
                            "type": "Action.OpenUrl",
                            "title": "View",
                            "url": url
                        }
                    ]
                }
            }
        ]
    }

    try {
        const response = await axios.post(webhookURL, card, {
            headers: {
                "content-type": "application/vnd.microsoft.card.adaptive"
            },
        });
        return `${response.status} - ${response.statusText}`;
    } catch (err) {
        return err;
    }
}
