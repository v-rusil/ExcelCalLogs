import { myConsole } from "./myConsole";

export async function callCognitive(jsonStringCDL:string):Promise<string>
{
    try {

        // Make a POST request to the sentiment analysis and entity recognition API (Azure Cognitive Services)
        //https://app.azure.com/72f988bf-86f1-41af-91ab-2d7cd011db47/subscriptions/01c7d5e3-40f7-4ba3-8bfb-78348f4db7db/resourceGroups/GroupNagome/providers/Microsoft.CognitiveServices/accounts/languagenagometest
        const apiUrl = "https://your-cognitive-services-endpoint.com/text/analytics/v3.1-preview.3/entities/recognition/general";
        const subscriptionKey = "YourSubscriptionKey";

        const response = await fetch(apiUrl, {
        method: "POST",
        headers: {
            "Content-Type": "application/json",
            "Ocp-Apim-Subscription-Key": subscriptionKey
        },
        body: JSON.stringify({
            documents: [
            {
                id: "1",
                text: jsonStringCDL
            }
            ]
        })
        });

        const entityResults = await response.json();

        // Extract the relevant information from the entity recognition results
        const timelineData = entityResults.documents[0].entities.map(entity => ({
        start: entity.text,
        end: entity.text,
        group: entity.category,
        content: entity.text,
        title: entity.category,
        description: entity.type
        }));



        return timelineData;
    } catch (error) {
        console.log(error);
        myConsole.log(error);
        return "";
    }
}