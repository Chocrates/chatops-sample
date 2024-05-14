const { ActivityHandler, MessageFactory, teamsGetChannelId, TeamsInfo, BotFrameworkAdapter } = require('botbuilder');

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

    async teamsCreateConversation(adapter, serviceUrl, message) {
        const teamsChannelId = "emulator"
        const conversationParameters = {
            isGroup: true,
            channelData: {
                channel: {
                    id: teamsChannelId
                }
            },
            activity: message,
            bot: {
                id: 'd1c86990-1146-11ef-b537-a7ce13724f54',
                name: 'Bot',
                role: 'bot'
            }
        };
        let conversationReference
        let newActivityId
        console.log('Right before sending the message')
        await adapter.createConversationAsync(
            'd1c86990-1146-11ef-b537-a7ce13724f54',
            teamsChannelId,
            serviceUrl,
            null,
            conversationParameters,
            async (turnContext) => {
                console.log('Inside the callback')
                conversationReference = TurnContext.getConversationReference(turnContext.activity);
                newActivityId = turnContext.activity.id;
            }
        )
        console.log('After sending the message')
        return [conversationReference, newActivityId]
    }
}

module.exports.EchoBot = EchoBot;
