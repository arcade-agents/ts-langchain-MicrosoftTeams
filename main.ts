"use strict";
import { getTools, confirm, arcade } from "./tools";
import { createAgent } from "langchain";
import {
  Command,
  MemorySaver,
  type Interrupt,
} from "@langchain/langgraph";
import chalk from "chalk";
import * as readline from "node:readline/promises";

// configure your own values to customize your agent

// The Arcade User ID identifies who is authorizing each service.
const arcadeUserID = process.env.ARCADE_USER_ID;
if (!arcadeUserID) {
  throw new Error("Missing ARCADE_USER_ID. Add it to your .env file.");
}
// This determines which MCP server is providing the tools, you can customize this to make a Slack agent, or Notion agent, etc.
// all tools from each of these MCP servers will be retrieved from arcade
const toolkits=['MicrosoftTeams'];
// This determines isolated tools that will be
const isolatedTools=[];
// This determines the maximum number of tool definitions Arcade will return
const toolLimit = 100;
// This prompt defines the behavior of the agent.
const systemPrompt = `# Introduction
Welcome to the Microsoft Teams AI Agent! This agent is designed to help users effectively interact with Microsoft Teams through various functionalities such as creating chats, retrieving messages, and managing teams and channels. The AI behaves in a ReAct architecture, allowing it to respond dynamically based on the user’s commands while accessing the appropriate tools seamlessly.

# Instructions
1. **Gather Information**: Begin by identifying what the user wants to accomplish. If necessary, ask clarifying questions.
2. **Execute Workflows**: Based on the user's requirements, follow one of the predefined workflows using the appropriate tools.
3. **Provide Feedback**: Once the actions are completed, inform the user of the results or provide the requested information in a clear format.
4. **Error Handling**: If any action fails or if required information is missing, provide helpful responses to guide the user on what to do next.

# Workflows

## Workflow 1: Create a Chat
- **Step 1**: Use `MicrosoftTeams_ListChats` to check if a chat already exists with the specified users.
- **Step 2**: If the chat exists, inform the user about it. If not, use `MicrosoftTeams_CreateChat` to create a new chat.

## Workflow 2: Retrieve Chat Messages
- **Step 1**: Use `MicrosoftTeams_GetChatMetadata` to get metadata about the chat.
- **Step 2**: Use `MicrosoftTeams_GetChatMessages` to retrieve messages from the chat.
  
## Workflow 3: Send a Message to a Chat
- **Step 1**: Verify if the chat exists using `MicrosoftTeams_ListChats`.
- **Step 2**: Use `MicrosoftTeams_SendMessageToChat` to send the specified message to the identified chat.

## Workflow 4: List Team Members
- **Step 1**: Use `MicrosoftTeams_ListTeams` to identify the relevant team.
- **Step 2**: Use `MicrosoftTeams_ListTeamMembers` to retrieve members of the selected team.

## Workflow 5: Search for Messages in Chats and Channels
- **Step 1**: Use `MicrosoftTeams_SearchMessages` to find messages that match the keywords specified by the user.

## Workflow 6: Manage Channels
- **Step 1**: Use `MicrosoftTeams_ListChannels` to list all channels of a team.
- **Step 2**: If the user wants to search, use `MicrosoftTeams_SearchChannels` to find specific channels based on keywords.

## Workflow 7: Reply to a Chat Message
- **Step 1**: Retrieve the chat message using `MicrosoftTeams_GetChatMessageById`.
- **Step 2**: Use `MicrosoftTeams_ReplyToChatMessage` to send the reply to the specified message.

By following these workflows, the AI agent will be able to help users efficiently navigate and utilize Microsoft Teams functionalities.`;
// This determines which LLM will be used inside the agent
const agentModel = process.env.OPENAI_MODEL;
if (!agentModel) {
  throw new Error("Missing OPENAI_MODEL. Add it to your .env file.");
}
// This allows LangChain to retain the context of the session
const threadID = "1";

const tools = await getTools({
  arcade,
  toolkits: toolkits,
  tools: isolatedTools,
  userId: arcadeUserID,
  limit: toolLimit,
});



async function handleInterrupt(
  interrupt: Interrupt,
  rl: readline.Interface
): Promise<{ authorized: boolean }> {
  const value = interrupt.value;
  const authorization_required = value.authorization_required;
  const hitl_required = value.hitl_required;
  if (authorization_required) {
    const tool_name = value.tool_name;
    const authorization_response = value.authorization_response;
    console.log("⚙️: Authorization required for tool call", tool_name);
    console.log(
      "⚙️: Please authorize in your browser",
      authorization_response.url
    );
    console.log("⚙️: Waiting for you to complete authorization...");
    try {
      await arcade.auth.waitForCompletion(authorization_response.id);
      console.log("⚙️: Authorization granted. Resuming execution...");
      return { authorized: true };
    } catch (error) {
      console.error("⚙️: Error waiting for authorization to complete:", error);
      return { authorized: false };
    }
  } else if (hitl_required) {
    console.log("⚙️: Human in the loop required for tool call", value.tool_name);
    console.log("⚙️: Please approve the tool call", value.input);
    const approved = await confirm("Do you approve this tool call?", rl);
    return { authorized: approved };
  }
  return { authorized: false };
}

const agent = createAgent({
  systemPrompt: systemPrompt,
  model: agentModel,
  tools: tools,
  checkpointer: new MemorySaver(),
});

async function streamAgent(
  agent: any,
  input: any,
  config: any
): Promise<Interrupt[]> {
  const stream = await agent.stream(input, {
    ...config,
    streamMode: "updates",
  });
  const interrupts: Interrupt[] = [];

  for await (const chunk of stream) {
    if (chunk.__interrupt__) {
      interrupts.push(...(chunk.__interrupt__ as Interrupt[]));
      continue;
    }
    for (const update of Object.values(chunk)) {
      for (const msg of (update as any)?.messages ?? []) {
        console.log("🤖: ", msg.toFormattedString());
      }
    }
  }

  return interrupts;
}

async function main() {
  const config = { configurable: { thread_id: threadID } };
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });

  console.log(chalk.green("Welcome to the chatbot! Type 'exit' to quit."));
  while (true) {
    const input = await rl.question("> ");
    if (input.toLowerCase() === "exit") {
      break;
    }
    rl.pause();

    try {
      let agentInput: any = {
        messages: [{ role: "user", content: input }],
      };

      // Loop until no more interrupts
      while (true) {
        const interrupts = await streamAgent(agent, agentInput, config);

        if (interrupts.length === 0) {
          break; // No more interrupts, we're done
        }

        // Handle all interrupts
        const decisions: any[] = [];
        for (const interrupt of interrupts) {
          decisions.push(await handleInterrupt(interrupt, rl));
        }

        // Resume with decisions, then loop to check for more interrupts
        // Pass single decision directly, or array for multiple interrupts
        agentInput = new Command({ resume: decisions.length === 1 ? decisions[0] : decisions });
      }
    } catch (error) {
      console.error(error);
    }

    rl.resume();
  }
  console.log(chalk.red("👋 Bye..."));
  process.exit(0);
}

// Run the main function
main().catch((err) => console.error(err));