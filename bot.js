const { TurnContext, ActivityHandler, MessageFactory, teamsGetChannelId, TeamsInfo, BotFrameworkAdapter } = require('botbuilder');


class EchoBot extends ActivityHandler {
    constructor() {
        super();
        // See https://aka.ms/about-bot-activity-message to learn more about the message and other activity types.
        this.onMessage(async (context, next) => {
            const replyText = `Echo: ${context.activity.text}`;
            await context.sendActivity(MessageFactory.text(replyText, replyText));
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });

        this.onMembersAdded(async (context, next) => {
            const membersAdded = context.activity.membersAdded;
            const welcomeText = 'Hello and welcome!';
            for (let cnt = 0; cnt < membersAdded.length; ++cnt) {
                if (membersAdded[cnt].id !== context.activity.recipient.id) {
                    await context.sendActivity(MessageFactory.text(welcomeText, welcomeText));
                }
            }
            // By calling next() you ensure that the next BotHandler is run.
            await next();
        });
    }

    async proactiveMessage(context) {
        const teamsChannelId = 'emulator'
        const activity = MessageFactory.text('This is a test message to a channel')
        const [reference] = await TeamsInfo.sendMessageToTeamsChannel(context, activity, teamsChannelId, undefined);//process.env.MicrosoftAppId);
    }

    async teamsCreateConversation(adapter) {
        const channelId = 'snip'
        const teamId = 'snip'
            
        const reference = TurnContext.getConversationReference({
            bot: {
                id: process.env.MicrosoftAppId,
                name: 'Your Bot'
            },
            channelId: channelId,
            conversation: {
                isGroup: true,
                conversationType: 'channel',
                id: `${teamId};messageid=${channelId}`
            },
            serviceUrl: 'https://smba.trafficmanager.net/amer/',
            user: {
                id: 'user-id-placeholder',
                name: 'User'
            }
        });

        await adapter.continueConversationAsync(reference, async (turnContext) => {
            await turnContext.sendActivity('Hello, this is a proactive message from the bot!');
        });

        console.log('After sending the message')
    }
}

module.exports.EchoBot = EchoBot;
