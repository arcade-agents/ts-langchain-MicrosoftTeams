---
title: "Build a MicrosoftTeams agent with LangChain (TypeScript) and Arcade"
slug: "ts-langchain-MicrosoftTeams"
framework: "langchain-ts"
language: "typescript"
toolkits: ["MicrosoftTeams"]
tools: []
difficulty: "beginner"
generated_at: "2026-03-12T01:35:06Z"
source_template: "ts_langchain"
agent_repo: ""
tags:
  - "langchain"
  - "typescript"
  - "microsoftteams"
---

# Build a MicrosoftTeams agent with LangChain (TypeScript) and Arcade

In this tutorial you'll build an AI agent using [LangChain](https://js.langchain.com/) with [LangGraph](https://langchain-ai.github.io/langgraphjs/) in TypeScript and [Arcade](https://arcade.dev) that can interact with MicrosoftTeams tools — with built-in authorization and human-in-the-loop support.

## Prerequisites

- The [Bun](https://bun.com) runtime
- An [Arcade](https://arcade.dev) account and API key
- An OpenAI API key

## Project Setup

First, create a directory for this project, and install all the required dependencies:

````bash
mkdir microsoftteams-agent && cd microsoftteams-agent
bun install @arcadeai/arcadejs @langchain/langgraph @langchain/core langchain chalk
````

## Start the agent script

Create a `main.ts` script, and import all the packages and libraries. Imports from 
the `"./tools"` package may give errors in your IDE now, but don't worry about those
for now, you will write that helper package later.

````typescript
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
````

## Configuration

In `main.ts`, configure your agent's toolkits, system prompt, and model. Notice
how the system prompt tells the agent how to navigate different scenarios and
how to combine tool usage in specific ways. This prompt engineering is important
to build effective agents. In fact, the more agentic your application, the more
relevant the system prompt to truly make the agent useful and effective at
using the tools at its disposal.

````typescript
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
const systemPrompt = "# ReAct Agent Prompt \u2014 Microsoft Teams Assistant\n\n## Introduction\nYou are a ReAct-style AI agent whose purpose is to help users interact with Microsoft Teams programmatically by calling the available Teams tools. Your job is to understand the user\u0027s intent, decide which Teams API tool(s) to call, call them with the correct parameters, handle errors/ambiguity, and return clear user-facing results or follow-up questions.\n\n## Instructions (how you should operate)\n- Follow the ReAct loop: Thought (internal), Action (call a tool), Observation (tool output), Thought, Action, ... then produce a Final Answer for the user.\n  - Use a Thought line to record your internal reasoning, but DO NOT reveal chain-of-thought content in the final user-facing message. Keep internal thoughts private.\n  - Each Action must specify the tool name and the exact JSON-like arguments to pass.\n  - After each Action include an Observation summarizing the tool output (the actual tool output will be injected by the system).\n- Parameter preferences and limits:\n  - Prefer providing IDs when available: channel_id over channel_name, team_id over team_name, user_ids over user_names.\n  - Respect API limits: message-listing tools return up to 50 items (max=50). Team/user listing endpoints have documented maxima (e.g., ListTeamMembers max 999). These endpoints generally do not support pagination.\n  - Do not call ListTeams/ListChannels/ListUsers unnecessarily \u2014 only call them if you need the list to disambiguate or the user explicitly requests it. (This reduces unnecessary API calls.)\n- Ambiguity handling:\n  - If required information is missing or multiple resources match (e.g., multiple teams), ask a concise clarifying question rather than guessing.\n  - If a tool returns an error or multiple candidates (e.g., multiple teams), present the choices and ask the user to pick.\n- Efficiency \u0026 safety:\n  - Avoid unnecessary calls (e.g., do not call ListTeams just to find a team when the user specified team_name).\n  - Observe privacy and security: do not display or infer sensitive data beyond what the user requests.\n- Error handling:\n  - If a tool fails, capture the error, map it to a user-friendly explanation (and possible fix), and propose the next steps.\n- Output:\n  - Final user-facing output should be concise, actionable, and not contain internal Thoughts.\n  - If you performed actions (messages sent, replies posted, chats created), state what you did and include IDs or names where useful.\n\n## ReAct Format Template (required)\nWhen you produce the reasoning/calls, follow this structure:\n\n```\nThought: [internal reasoning \u2014 DO NOT reveal in final answer]\nAction: \u003cToolName\u003e \u003cJSON arguments\u003e\nObservation: \u003ctool output (system will provide)\u003e\n... (repeat)\nFinal Answer: \u003cuser-facing summary / next question\u003e\n```\n\n## Workflows\nBelow are typical workflows you will follow. Each workflow lists the recommended sequence of tools and notes about parameter selection and edge cases. Use these as patterns; adapt when the user\u2019s request requires deviations.\n\n1) Send a message to a channel\n- Goal: Post a new message to a channel.\n- Preferred sequence:\n  - Action: MicrosoftTeams_SendMessageToChannel { \"message\": \"...\", \"channel_id_or_name\": \"\u003cchannel-id-or-name\u003e\", \"team_id_or_name\": \"\u003cteam-id-or-name (optional)\u003e\" }\n- Notes:\n  - Prefer channel_id and team_id if available.\n  - If the user only gave a channel name and you belong to multiple teams, ask which team unless the tool can uniquely resolve it.\n\nExample:\n```\nThought: Post project update to #general in Team A\nAction: MicrosoftTeams_SendMessageToChannel {\"message\":\"Project update: ...\",\"channel_id_or_name\":\"general\",\"team_id_or_name\":\"Team A\"}\nObservation: { ... }  // tool output injected\nFinal Answer: Message posted to #general in Team A (message id: ...).\n```\n\n2) Reply to a channel message\n- Goal: Reply to a specific channel message.\n- Preferred sequence:\n  - If you have the message_id and channel: MicrosoftTeams_ReplyToChannelMessage { \"reply_content\": \"...\", \"message_id\": \"\u003cid\u003e\", \"channel_id_or_name\": \"\u003cchannel-id-or-name\u003e\", \"team_id_or_name\": \"\u003cteam-id-or-name (optional)\u003e\" }\n  - If you don\u2019t have message_id: Get channel messages first -\u003e identify message -\u003e call ReplyToChannelMessage.\n    - MicrosoftTeams_GetChannelMessages { \"channel_id\": \"\u003cid\u003e\", \"team_id_or_name\": \"\u003cteam\u003e\", \"limit\": N }\n    - Then MicrosoftTeams_ReplyToChannelMessage as above.\n- Notes:\n  - GetChannelMessages returns newest-first and max 50 messages. If the target message is old, ask for a date-range or more identifying info.\n\nExample:\n```\nThought: Need to reply to the message with id 123\nAction: MicrosoftTeams_ReplyToChannelMessage {\"reply_content\":\"Thanks \u2014 noted.\",\"message_id\":\"123\",\"channel_id_or_name\":\"general\"}\nObservation: { ... }\nFinal Answer: Replied to the message in #general.\n```\n\n3) Read messages from a channel\n- Goal: Fetch messages from a channel for review, optionally filtering by limit.\n- Sequence:\n  - MicrosoftTeams_GetChannelMessages { \"channel_id\": \"\u003cid\u003e\", \"team_id_or_name\": \"\u003cteam\u003e\", \"limit\": N }\n- Notes:\n  - No pagination; limit \u2264 50. If the user needs older messages beyond 50, ask for a date range or narrower filters.\n\n4) Get replies to a specific channel message\n- Goal: Retrieve threaded replies under a channel message.\n- Sequence:\n  - MicrosoftTeams_GetChannelMessageReplies { \"message_id\": \"\u003cid\u003e\", \"channel_id_or_name\": \"\u003cchannel-id-or-name\u003e\", \"team_id_or_name\": \"\u003cteam-id-or-name (optional)\u003e\" }\n- Notes:\n  - Provide the message_id and channel. If not available, find the message via GetChannelMessages first.\n\n5) Create a group chat and send a message\n- Goal: Start a new chat (or get existing one) with specified members, then send a message.\n- Sequence:\n  - MicrosoftTeams_CreateChat { \"user_ids\": [...], \"user_names\": [...] }\n    - The API will return existing chat if it already exists.\n  - MicrosoftTeams_SendMessageToChat { \"chat_id\": \"\u003cchat-id-from-create\u003e\", \"message\":\"...\" }\n  - Alternatively, if you have user_ids but want to avoid CreateChat, you may call SendMessageToChat with user_ids directly.\n- Notes:\n  - Prefer user_ids. Max 20 users when using GetChatMetadata-related calls; check tool docs for other limits.\n\nExample:\n```\nThought: Create chat with Alice and Bob, send kickoff note\nAction: MicrosoftTeams_CreateChat {\"user_names\":[\"Alice Smith\",\"Bob Jones\"]}\nObservation: { \"chat_id\": \"chat-456\", ... }\nAction: MicrosoftTeams_SendMessageToChat {\"chat_id\":\"chat-456\",\"message\":\"Kickoff meeting at 10am\"}\nObservation: { ... }\nFinal Answer: Chat created/located and message sent (chat id: chat-456).\n```\n\n6) Send a message to an existing chat or reply to a chat message\n- Goal A (send): MicrosoftTeams_SendMessageToChat { \"chat_id\": \"\u003cid\u003e\" } or { \"user_ids\": [...] / \"user_names\": [...] }\n- Goal B (reply): MicrosoftTeams_ReplyToChatMessage { \"reply_content\":\"...\", \"message_id\":\"...\", \"chat_id\":\"\u003cid\u003e\" } or provide user_ids/user_names in place of chat_id.\n- Notes:\n  - Exactly one of chat_id OR user_ids/user_names must be provided in SendMessageToChat calls.\n  - For ReplyToChatMessage, provide chat_id or user_ids/user_names exactly as required.\n\n7) Retrieve messages from a chat\n- Goal: Fetch chat messages with optional date filters.\n- Sequence:\n  - MicrosoftTeams_GetChatMessages { \"chat_id\": \"\u003cid\u003e\" } OR { \"user_ids\": [...]/\"user_names\": [...], \"start_datetime\":\"YYYY-MM-DD\", \"end_datetime\":\"YYYY-MM-DD\", \"limit\": N }\n- Notes:\n  - Messages sorted descending by created_datetime, max 50, no pagination.\n\n8) Search for messages, channels, teams, or people, then act\n- Use appropriate search endpoints:\n  - SearchMessages { \"keywords\": \"...\" }\n  - SearchChannels { \"keywords\": [\"keyword\"] , \"team_id_or_name\": \"...\" }\n  - SearchUsers { \"keywords\": [\"name\"] } (tenant users)\n  - SearchPeople { \"keywords\": [\"name\"] } (people the user has interacted with; includes external)\n  - SearchTeams { \"team_name_starts_with\": \"...\" }\n- Typical sequence:\n  - Search -\u003e review results -\u003e call a targeted action (GetChannelMessages, GetChatMessages, SendMessageToChannel, CreateChat, etc.)\n- Notes:\n  - Search APIs may be eventually consistent: recent items might not appear immediately.\n\n9) Get team or channel metadata and members\n- Goal: Inspect channel/team members or metadata before taking action.\n- Tools:\n  - MicrosoftTeams_GetChannelMetadata { \"channel_id\": \"\u003cid\u003e\" } or { \"channel_name\": \"\u003cname\u003e\", \"team_id_or_name\": \"\u003cteam\u003e\" }\n  - MicrosoftTeams_GetTeam { \"team_id\": \"\u003cid\u003e\" } or { \"team_name\": \"\u003cname\u003e\" }\n  - MicrosoftTeams_ListTeamMembers { \"team_id\": \"\u003cid\u003e\", \"limit\": N, \"offset\": M }\n  - MicrosoftTeams_ListChannels { \"team_id_or_name\": \"\u003cteam\u003e\" }\n- Notes:\n  - The API returns up to the first 999 members; prefer channel_id when available.\n\n10) Common cross-cutting rules \u0026 examples\n- Prefer specific identifiers:\n  - Use channel_id, team_id, chat_id, user_ids whenever possible.\n- Avoid redundant calls:\n  - Do not call MicrosoftTeams_ListTeams or other broad listing tools unless you need them to disambiguate or the user explicitly asked to \"list\" resources.\n- When you must handle multiple matches:\n  - Present a short numbered list of candidates and ask the user to choose (e.g., \"I found these 3 teams matching \u0027Acme\u0027 \u2014 which one did you mean? 1) Acme Corp (id: t1), 2) Acme Dev (id: t2)...\")\n\nExample multi-step interaction (search user, create chat, send message):\n```\nThought: Find user \"Jane Doe\" then start chat and send intro\nAction: MicrosoftTeams_SearchUsers {\"keywords\":[\"Jane Doe\"], \"limit\": 10}\nObservation: { ... } // contains user id(s)\nThought: Use Jane\u0027s id to create chat\nAction: MicrosoftTeams_CreateChat {\"user_ids\":[\"user-123\"]}\nObservation: { \"chat_id\":\"chat-789\", ... }\nAction: MicrosoftTeams_SendMessageToChat {\"chat_id\":\"chat-789\", \"message\":\"Hi Jane, I\u0027m connecting you with the team.\"}\nObservation: { ... }\nFinal Answer: Chat created and message sent to Jane (chat id: chat-789).\n```\n\n## Error \u0026 edge-case handling examples\n- If GetChannelMessages returns no results but user expects messages, ask:\n  - \"I found no messages in that channel. Do you want me to increase the time range or check another channel?\"\n- If a tool returns multiple teams or channels, reply with:\n  - \"I found multiple matches; please pick one: 1) Team A (id: t1), 2) Team B (id: t2).\"\n- If the user requests an action that would exceed API limits (e.g., \u003e50 messages), state the limit and request additional constraints (date range, keywords).\n\n---\n\nUse the workflows above as canonical patterns. Always remember: keep internal Thoughts private, prefer precise identifiers, minimize unnecessary calls, and ask clarifying questions when in doubt.";
// This determines which LLM will be used inside the agent
const agentModel = process.env.OPENAI_MODEL;
if (!agentModel) {
  throw new Error("Missing OPENAI_MODEL. Add it to your .env file.");
}
// This allows LangChain to retain the context of the session
const threadID = "1";
````

Set the following environment variables in a `.env` file:

````bash
ARCADE_API_KEY=your-arcade-api-key
ARCADE_USER_ID=your-arcade-user-id
OPENAI_API_KEY=your-openai-api-key
OPENAI_MODEL=gpt-5-mini
````

## Implementing the `tools.ts` module

The `tools.ts` module fetches Arcade tool definitions and converts them to LangChain-compatible tools using Arcade's Zod schema conversion:

### Create the file and import the dependencies

Create a `tools.ts` file, and add import the following. These will allow you to build the helper functions needed to convert Arcade tool definitions into a format that LangChain can execute. Here, you also define which tools will require human-in-the-loop confirmation. This is very useful for tools that may have dangerous or undesired side-effects if the LLM hallucinates the values in the parameters. You will implement the helper functions to require human approval in this module.

````typescript
import { Arcade } from "@arcadeai/arcadejs";
import {
  type ToolExecuteFunctionFactoryInput,
  type ZodTool,
  executeZodTool,
  isAuthorizationRequiredError,
  toZod,
} from "@arcadeai/arcadejs/lib/index";
import { type ToolExecuteFunction } from "@arcadeai/arcadejs/lib/zod/types";
import { tool } from "langchain";
import {
  interrupt,
} from "@langchain/langgraph";
import readline from "node:readline/promises";

// This determines which tools require human in the loop approval to run
const TOOLS_WITH_APPROVAL = ['MicrosoftTeams_CreateChat', 'MicrosoftTeams_ReplyToChannelMessage', 'MicrosoftTeams_ReplyToChatMessage', 'MicrosoftTeams_SendMessageToChannel', 'MicrosoftTeams_SendMessageToChat'];
````

### Create a confirmation helper for human in the loop

The first helper that you will write is the `confirm` function, which asks a yes or no question to the user, and returns `true` if theuser replied with `"yes"` and `false` otherwise.

````typescript
// Prompt user for yes/no confirmation
export async function confirm(question: string, rl?: readline.Interface): Promise<boolean> {
  let shouldClose = false;
  let interface_ = rl;

  if (!interface_) {
      interface_ = readline.createInterface({
          input: process.stdin,
          output: process.stdout,
      });
      shouldClose = true;
  }

  const answer = await interface_.question(`${question} (y/n): `);

  if (shouldClose) {
      interface_.close();
  }

  return ["y", "yes"].includes(answer.trim().toLowerCase());
}
````

Tools that require authorization trigger a LangGraph interrupt, which pauses execution until the user completes authorization in their browser.

### Create the execution helper

This is a wrapper around the `executeZodTool` function. Before you execute the tool, however, there are two logical checks to be made:

1. First, if the tool the agent wants to invoke is included in the `TOOLS_WITH_APPROVAL` variable, human-in-the-loop is enforced by calling `interrupt` and passing the necessary data to call the `confirm` helper. LangChain will surface that `interrupt` to the agentic loop, and you will be required to "resolve" the interrupt later on. For now, you can assume that the reponse of the `interrupt` will have enough information to decide whether to execute the tool or not, depending on the human's reponse.
2. Second, if the tool was approved by the human, but it doesn't have the authorization of the integration to run, then you need to present an URL to the user so they can authorize the OAuth flow for this operation. For this, an execution is attempted, that may fail to run if the user is not authorized. When it fails, you interrupt the flow and send the authorization request for the harness to handle. If the user authorizes the tool, the harness will reply with an `{authorized: true}` object, and the system will retry the tool call without interrupting the flow.

````typescript
export function executeOrInterruptTool({
  zodToolSchema,
  toolDefinition,
  client,
  userId,
}: ToolExecuteFunctionFactoryInput): ToolExecuteFunction<any> {
  const { name: toolName } = zodToolSchema;

  return async (input: unknown) => {
    try {

      // If the tool is on the list that enforces human in the loop, we interrupt the flow and ask the user to authorize the tool

      if (TOOLS_WITH_APPROVAL.includes(toolName)) {
        const hitl_response = interrupt({
          authorization_required: false,
          hitl_required: true,
          tool_name: toolName,
          input: input,
        });

        if (!hitl_response.authorized) {
          // If the user didn't approve the tool call, we throw an error, which will be handled by LangChain
          throw new Error(
            `Human in the loop required for tool call ${toolName}, but user didn't approve.`
          );
        }
      }

      // Try to execute the tool
      const result = await executeZodTool({
        zodToolSchema,
        toolDefinition,
        client,
        userId,
      })(input);
      return result;
    } catch (error) {
      // If the tool requires authorization, we interrupt the flow and ask the user to authorize the tool
      if (error instanceof Error && isAuthorizationRequiredError(error)) {
        const response = await client.tools.authorize({
          tool_name: toolName,
          user_id: userId,
        });

        // We interrupt the flow here, and pass everything the handler needs to get the user's authorization
        const interrupt_response = interrupt({
          authorization_required: true,
          authorization_response: response,
          tool_name: toolName,
          url: response.url ?? "",
        });

        // If the user authorized the tool, we retry the tool call without interrupting the flow
        if (interrupt_response.authorized) {
          const result = await executeZodTool({
            zodToolSchema,
            toolDefinition,
            client,
            userId,
          })(input);
          return result;
        } else {
          // If the user didn't authorize the tool, we throw an error, which will be handled by LangChain
          throw new Error(
            `Authorization required for tool call ${toolName}, but user didn't authorize.`
          );
        }
      }
      throw error;
    }
  };
}
````

### Create the tool retrieval helper

The last helper function of this module is the `getTools` helper. This function will take the configurations you defined in the `main.ts` file, and retrieve all of the configured tool definitions from Arcade. Those definitions will then be converted to LangGraph `Function` tools, and will be returned in a format that LangChain can present to the LLM so it can use the tools and pass the arguments correctly. You will pass the `executeOrInterruptTool` helper you wrote in the previous section so all the bindings to the human-in-the-loop and auth handling are programmed when LancChain invokes a tool.


````typescript
// Initialize the Arcade client
export const arcade = new Arcade();

export type GetToolsProps = {
  arcade: Arcade;
  toolkits?: string[];
  tools?: string[];
  userId: string;
  limit?: number;
}


export async function getTools({
  arcade,
  toolkits = [],
  tools = [],
  userId,
  limit = 100,
}: GetToolsProps) {

  if (toolkits.length === 0 && tools.length === 0) {
      throw new Error("At least one tool or toolkit must be provided");
  }

  // Todo(Mateo): Add pagination support
  const from_toolkits = await Promise.all(toolkits.map(async (tkitName) => {
      const definitions = await arcade.tools.list({
          toolkit: tkitName,
          limit: limit
      });
      return definitions.items;
  }));

  const from_tools = await Promise.all(tools.map(async (toolName) => {
      return await arcade.tools.get(toolName);
  }));

  const all_tools = [...from_toolkits.flat(), ...from_tools];
  const unique_tools = Array.from(
      new Map(all_tools.map(tool => [tool.qualified_name, tool])).values()
  );

  const arcadeTools = toZod({
    tools: unique_tools,
    client: arcade,
    executeFactory: executeOrInterruptTool,
    userId: userId,
  });

  // Convert Arcade tools to LangGraph tools
  const langchainTools = arcadeTools.map(({ name, description, execute, parameters }) =>
    (tool as Function)(execute, {
      name,
      description,
      schema: parameters,
    })
  );

  return langchainTools;
}
````

## Building the Agent

Back on the `main.ts` file, you can now call the helper functions you wrote to build the agent.

### Retrieve the configured tools

Use the `getTools` helper you wrote to retrieve the tools from Arcade in LangChain format:

````typescript
const tools = await getTools({
  arcade,
  toolkits: toolkits,
  tools: isolatedTools,
  userId: arcadeUserID,
  limit: toolLimit,
});
````

### Write an interrupt handler

When LangChain is interrupted, it will emit an event in the stream that you will need to handle and resolve based on the user's behavior. For a human-in-the-loop interrupt, you will call the `confirm` helper you wrote earlier, and indicate to the harness whether the human approved the specific tool call or not. For an auth interrupt, you will present the OAuth URL to the user, and wait for them to finishe the OAuth dance before resolving the interrupt with `{authorized: true}` or `{authorized: false}` if an error occurred:

````typescript
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
````

### Create an Agent instance

Here you create the agent using the `createAgent` function. You pass the system prompt, the model, the tools, and the checkpointer. When the agent runs, it will automatically use the helper function you wrote earlier to handle tool calls and authorization requests.

````typescript
const agent = createAgent({
  systemPrompt: systemPrompt,
  model: agentModel,
  tools: tools,
  checkpointer: new MemorySaver(),
});
````

### Write the invoke helper

This last helper function handles the streaming of the agent’s response, and captures the interrupts. When the system detects an interrupt, it adds the interrupt to the `interrupts` array, and the flow interrupts. If there are no interrupts, it will just stream the agent’s to your console.

````typescript
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
````

### Write the main function

Finally, write the main function that will call the agent and handle the user input.

Here the `config` object configures the `thread_id`, which tells the agent to store the state of the conversation into that specific thread. Like any typical agent loop, you:

1. Capture the user input
2. Stream the agent's response
3. Handle any authorization interrupts
4. Resume the agent after authorization
5. Handle any errors
6. Exit the loop if the user wants to quit

````typescript
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
````

## Running the Agent

### Run the agent

```bash
bun run main.ts
```

You should see the agent responding to your prompts like any model, as well as handling any tool calls and authorization requests.

## Next Steps

- Clone the [repository](https://github.com/arcade-agents/ts-langchain-MicrosoftTeams) and run it
- Add more toolkits to the `toolkits` array to expand capabilities
- Customize the `systemPrompt` to specialize the agent's behavior
- Explore the [Arcade documentation](https://docs.arcade.dev) for available toolkits

