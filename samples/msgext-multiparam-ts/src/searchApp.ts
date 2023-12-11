import { AdaptiveCards } from "@microsoft/adaptivecards-tools";
import {
  TeamsActivityHandler,
  CardFactory,
  TurnContext,
  MessagingExtensionQuery,
  MessagingExtensionResponse,
} from "botbuilder";
import stockCard from "./cards/stock.json";
import stockCardData from "./cards/stock.data.json";
import { StockData } from "./types";
import config from "./config";
import {handleMessageExtensionQueryWithSSO, OnBehalfOfCredentialAuthConfig, MessageExtensionTokenResponse, OnBehalfOfUserCredential} from "@microsoft/teamsfx";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials";
import "isomorphic-fetch";

const oboAuthConfig: OnBehalfOfCredentialAuthConfig = {
  authorityHost: config.authorityHost,
  clientId: config.clientId,
  tenantId: config.tenantId,
  clientSecret: config.clientSecret,
};

const initialLoginEndpoint = `https://${config.botDomain}/auth-start.html`;

export class SearchApp extends TeamsActivityHandler {

  public async handleTeamsMessagingExtensionQuery(
    context: TurnContext,
    query: MessagingExtensionQuery
  ): Promise<any> {

    return await handleMessageExtensionQueryWithSSO(
      context,
      oboAuthConfig,
      initialLoginEndpoint,
      ["User.Read"],
      async (token : MessageExtensionTokenResponse) => {
        
        // demo code to call Microsoft Graph API with the token
        const credential = new OnBehalfOfUserCredential(
          token.ssoToken,
          oboAuthConfig
        );
        const authProvider = new TokenCredentialAuthenticationProvider(
          credential,
          {
            scopes: ["User.Read"],
          }
        );
        const graphClient = Client.initWithMiddleware({
          authProvider: authProvider,
        });
        const profile = await graphClient.api("/me").get();
        console.log(`Your profile from Microsoft Graph is ${JSON.stringify(profile)}`);

        const { parameters } = query;

        // Basic: Find stocks in NASDAQ Stocks
        // [
        //   { name: 'StockIndex', value: 'NASDAQ' },
        //   { name: 'NumberofStocks', value: '' },
        //   { name: 'P/B', value: '' },
        //   { name: 'P/E', value: '' }
        // ]

        // Advanced: Find top 10 stocks in NASDAQ Stocks with P/B < 2 and P/E < 30
        // [
        //   { name: 'StockIndex', value: '' },
        //   { name: 'NumberofStocks', value: 'Top:10' },
        //   { name: 'P/B', value: '<2' },
        //   { name: 'P/E', value: '<30' }
        // ]

        const stockIndex = parameters[0].value;
        const numberOfStocks = parameters[1].value;
        const pB = parameters[2].value;
        const pE = parameters[3].value;

        console.log(
          `🔍 StockIndex: '${stockIndex}' | NumberofStocks: '${numberOfStocks}' | P/B: '${pB}' | P/E: '${pE}'`
        );

        const attachments = stockCardData.map((stock) => {
          const card = AdaptiveCards.declare<StockData>(stockCard).render(stock);
          const resultCard = CardFactory.adaptiveCard(card);
          const previewCard = CardFactory.heroCard(stock.companyName, stock.symbol);
          return { ...resultCard, previewCard };
        });

        return {
          composeExtension: {
            type: "result",
            attachmentLayout: "list",
            attachments,
          },
        };
      }
    );
  }
}
