import { Client } from "@microsoft/microsoft-graph-client";
import { ToolUtility } from "@azure/ai-agents";

export class FunctionToolExecutor {

    private functionTools;
    private graphClient: Client;

    constructor(graphClient: Client) {

        this.getCopilotData = this.getCopilotData.bind(this);

        this.graphClient = graphClient;

        this.functionTools = [
            {
                func: this.getCopilotData,
                ...ToolUtility.createFunctionTool({
                    name: "getCopilotData",
                    description: "Retrieve content for current user",
                    type: "function",
                    parameters: {
                        type: "object",
                        properties: {
                            query: { type: "string", description: "The user original and non modified query" },
                        }
                    }
                } as any)
            }
        ];
    }

    public async getCopilotData(query: string) {
        
        const response = await this.graphClient.api("/copilot/retrieval").post({
            queryString: query,
            dataSource: "sharepoint",
            resourceMetadata: ['title','webUrl'],
            maximumNumberOfResults: 5
        });

        return response.retrievalHits;
    }

    public async invokeTool(toolCall) {
      console.log(`Function tool call - ${toolCall.function.name}`);
      const args = [];
      if (toolCall.function.arguments) {
        try {
          const params = JSON.parse(toolCall.function.arguments);
          for (const key in params) {
            if (Object.prototype.hasOwnProperty.call(params, key)) {
              args.push(params[key]);
            }
          }
        } catch (error) {
          console.error(`Failed to parse parameters: ${toolCall.function.arguments}`, error);
          return undefined;
        }
      }
      const result = await this.functionTools
        .find((tool) => tool.definition.function.name === toolCall.function.name)
        ?.func(...args);
      return result
        ? {
            toolCallId: toolCall.id,
            output: JSON.stringify(result),
          }
        : undefined;
    }

    public getFunctionDefinitions() {
      return this.functionTools.map((tool) => {
        return tool.definition;
      });
    }
}