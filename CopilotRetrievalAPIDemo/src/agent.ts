import { ClientSecretCredential } from "@azure/identity";
import { ActivityTypes } from "@microsoft/agents-activity";
import { AgentApplication, Citation, MemoryStorage, MessageFactory, TurnContext, TurnState } from "@microsoft/agents-hosting";
import { isOutputOfType } from "@azure/ai-agents";
import { AIProjectClient  } from "@azure/ai-projects";
import { Client } from "@microsoft/microsoft-graph-client";
import { FunctionToolExecutor } from "./toolexecutor";
import { Utils } from "./utils";

export class CustomAgent extends AgentApplication<TurnState> {

  constructor () {
    super({
      storage: new MemoryStorage(),    
      authorization: {
        graph: { 
          text: 'Sign in with Microsoft Graph', 
          title: 'Graph Sign In'
        },
      }
    })

    this._onMessage = this._onMessage.bind(this);
    this._singinSuccess = this._singinSuccess.bind(this);
    this._singinFailure = this._singinFailure.bind(this);

    this.authorization.onSignInSuccess(this._singinSuccess);
    this.authorization.onSignInFailure(this._singinFailure);

    this.onActivity(ActivityTypes.Message, this._onMessage, ['graph']);
  }

  private async _onMessage(context: TurnContext, state: TurnState) {

    try {

        let userTokenResponse;
        userTokenResponse = await this.authorization.getToken(context, 'graph');

        if (userTokenResponse && userTokenResponse?.token) {    

          context.streamingResponse.setGeneratedByAILabel(true)
          await context.streamingResponse.queueInformativeUpdate('Crafting your answer...')
          await this._invokeAgent(context, userTokenResponse?.token);
        }
      
    } catch (ex) {
      await context.sendActivity(`On message error. Details: ${JSON.stringify(ex)}`);       
    }
  }; 

  private async _singinSuccess(context: TurnContext, state: TurnState, authId?: string): Promise<void> {
    await context.sendActivity(MessageFactory.text(`User signed in successfully in ${authId}`))
  }

  private async _singinFailure (context: TurnContext, state: TurnState, authId?: string, err?: string): Promise<void> {
    await context.sendActivity(MessageFactory.text(`Signing Failure in auth handler: ${authId} with error: ${err}`))
  }

  private async _invokeAgent(context: TurnContext, token: string) {
  
    const agentId = process.env["ENV_AZURE_DEPLOY_AGENT_ID"];
    const aiFoundryProjectEndpoint = process.env["ENV_AZURE_DEPLOY_AI_FOUNDRY_PROJECT_ENDPOINT"];
    
    const credential = new ClientSecretCredential(
        process.env["ENV_AZURE_APP_TENANT_ID"],
        process.env["ENV_AZURE_APP_CLIENT_ID"],
        process.env["ENV_AZURE_APP_CLIENT_SECRET"],
    );   

    const project = new AIProjectClient(aiFoundryProjectEndpoint, credential);
    const agent = await project.agents.getAgent(agentId);
    const thread = await project.agents.threads.create();
    const message = await project.agents.messages.create(thread.id, "user", context.activity.text);
    
    let graphClient;
    if (token) {
      graphClient = Client.initWithMiddleware({
          authProvider: {
              getAccessToken: async () => {
                  return token
              },
          }
      });
    }

    const functionToolExecutor = new FunctionToolExecutor(graphClient);
    const functionTools = functionToolExecutor.getFunctionDefinitions();
    
    if (agent) {
        // Register custom tool
        await project.agents.updateAgent(agentId, {
            tools: functionTools
        });
    }

    console.log(`Created message, ID: ${message.id}`);

    let run = await project.agents.runs.create(thread.id, agent.id);
    let toolResponseOutput: any[] = [];

    // Wait for agent processing
    while (run.status === "queued" || run.status === "in_progress") {
      await new Promise((resolve) => setTimeout(resolve, 1000));
      run = await project.agents.runs.get(thread.id, run.id);

      // Determine if a tool should be executed
      if (run.status === "requires_action" && run.requiredAction) {

        console.log(`Run requires action - ${run.requiredAction}`);

        if (isOutputOfType(run.requiredAction, "submit_tool_outputs")) {
            
            const submitToolOutputsActionOutput = run.requiredAction;
            const toolCalls = submitToolOutputsActionOutput['submitToolOutputs'].toolCalls;
            const toolResponses = [];

            for (const toolCall of toolCalls) {

                if (isOutputOfType(toolCall, "function")) {
                    const toolResponse = await functionToolExecutor.invokeTool(toolCall);
                    if (toolResponse) {

                        toolResponseOutput = JSON.parse(toolResponse.output);
                        toolResponses.push(toolResponse);
                    }
                }
            }

            if (toolResponses.length > 0) {
                run = await project.agents.runs.submitToolOutputs(thread.id, run.id, toolResponses);
                console.log(`Submitted tool response - ${run.status}`);
            }
        }
      }
    }

    if (run.status === "failed") {
        console.error(`Run failed: `, run.lastError);
    }

    console.log(`Run completed with status: ${run.status}`);

    // Retrieve messages
    const messages = await project.agents.messages.list(thread.id, { order: "asc"});

    // Display messages
    const threadMessages = [];
    for await (const m of messages) {
        const content = m.content.find((c) => c.type === "text" && "text" in c);
        if (content) {
            threadMessages.push(m)
        }
    }

    // Get the last message sent by the agent and out put it to the user
    const agentAnswer: string = threadMessages[threadMessages.length-1].content[0].text.value;
    
    const links = Utils.extractMarkdownLinks(agentAnswer);
    const streamingCitations = links.map((link, i) => {

      // Get extracts from the tool output
      const snippet = toolResponseOutput.find((hit) => hit.webUrl === decodeURIComponent(link.url))?.extracts[0]?.text; // Get the first extract matching that URL (basic example)
      const citation: Citation = {
        title: link.title,
        url: link.url,
        content: "",
        filepath: link.url
      };

      return citation;
    });

    await context.streamingResponse.setCitations(streamingCitations);    
    await context.streamingResponse.queueTextChunk(Utils.replaceMarkdownLinksWithOrder(agentAnswer));
    await context.streamingResponse.endStream();
  }       
}