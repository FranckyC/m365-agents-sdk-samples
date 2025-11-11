# Combine Copilot Retrieval API, M365 Agents SDK and AI Foundry Agent Service

This solution describes how to use the Microsoft 365 Agents SDK, AI Foundry Agent Service and the Copilot retrieval Graph API.

Follow the step-by-step instructions from this blog post to get started ➡️ [https://blog.franckcornu.com/post/copilot-retrieval-api-ai-foundry/](https://blog.franckcornu.com/post/copilot-retrieval-api-ai-foundry).

## Context


> **What is the Azure AI Foundry Agent Service?**

  The Azure AI Agent Service is a component of the broader Azure AI Foundry platform, which enables you to manage all your AI assets—such as models, data, and test workloads—within a unified workspace organized into projects. From a technical standpoint, the Agent Service acts as a wrapper around the OpenAI Assistants API, offering both a configuration UI (similar to Copilot Studio) and SDK options. Through these, you can define your agent's model (deployed within your AI Foundry resource), instructions, tools, and knowledge sources.

  Unlike Copilot Studio, however, the Agent Service does not integrate with any prebuilt user-facing interfaces (such as Teams or Microsoft Copilot). It is the developer's responsibility to create and integrate such interfaces using the available SDKs for Node.js, C#, or Python.

  In essence, the Agent Service provides a streamlined backend for building and managing AI agents without relying on external orchestration frameworks like [LangChain](https://www.langchain.com/)
  or [Microsoft Agent Framework](https://learn.microsoft.com/en-us/agent-framework/overview/agent-framework-overview)
  . It is purpose-built for enterprise-grade deployments, supporting capabilities such as private network isolation, governance, guardrails, and policy controls—ensuring secure and compliant integration within organizational infrastructure.

> **What is the Copilot Retrieval API?**

The Copilot Retrieval API is a new Microsoft Graph API that enables access to content from the Microsoft 365 semantic index — the same index used by Copilot for Microsoft 365 and Copilot Studio as their underlying knowledge source. Before this API, there was essentially no reliable or secure way to access data from this index, making it difficult to implement Retrieval-Augmented Generation (RAG) patterns using Microsoft 365 content. This limitation was particularly frustrating, as organizations couldn't fully leverage their own Microsoft 365 data for custom Copilot-based solutions, even with paid licenses. Alternative approaches, such as indexing Microsoft 365 data (e.g., SharePoint) into Azure AI Search, were far from ideal — they introduced additional costs, lower security, significant limitations, and heavy maintenance requirements.

The version 1 of the Copilot Retrieval API has now been officially released and is generally available, meaning it's safe and ready for use in production environments.

> **What is the Microsoft 365 SDK?**

The Microsoft 365 SDK is essentially the evolution of the Bot Framework v4, designed to simplify the development of custom agent solutions for Microsoft 365 — particularly for Teams and Copilot. Unlike frameworks such as the Microsoft Agent Framework or LangChain, the Microsoft 365 Agents SDK is not an LLM orchestration framework. Instead, it focuses on streamlining integration with Microsoft 365 services for agent implementations, while still relying on the Azure Bot Service behind the scenes.

As noted by Andrew Connell [here](https://www.linkedin.com/posts/andrewconnell_microsoft-keeps-shooting-itself-in-the-foot-activity-7394007889503371264-pY6P?utm_source=share&utm_medium=member_desktop&rcm=ACoAAAM4n1QBzMd5cpRFZG3Puhwjc4kk-YOS9Kc), the SDK is still very new, changes frequently, and currently lacks comprehensive documentation and guidance. Many of the available samples and examples overlap with older Bot Framework content, which often leads to confusion among developers — especially when dealing with common scenarios such as authentication handling.

> **Why combine these three?**

This combination is particularly effective for building enterprise-grade agents that meet the following criteria:

- Designed to be consumed through built-in Microsoft 365 channels such as Teams and Copilot.
- Able to leverage pre-indexed Microsoft 365 data from native sources (e.g., SharePoint, OneDrive) or line-of-business (LOB) systems via Graph connectors, ensuring compliance and security.
- Provide control over AI models, including fine-tuning capabilities.
- Offer granular control over testing processes and analytics.
- Support environmental consistency across development, UAT, and production through programmatic provisioning with Bicep templates.
- Enable private network isolation within Azure.
- Allow centralized management of a fleet of dedicated company agents through a shared infrastructure.
- Remain flexible and extensible, allowing the solution to evolve without running into hard limitations (e.g., data sources, tools, or connected agents).

From my recent experience, I adopted this approach as a more robust alternative to Copilot declarative agents, which currently lack several key enterprise features, such as:

- No efficient or automated way to measure RAG performance — it requires a manual and cumbersome process.
- Lack of analytics and insights for stakeholders, making it difficult to monitor agent usage.
- No agent-to-agent communication capabilities.
- Limited feedback control — for instance, thumbs-up/down feedback is sent to Microsoft, not your organization.
- Performance issues when using API plugins.

While declarative agents can still be useful in certain scenarios, their value proposition within the Microsoft 365 Copilot extensibility ecosystem now feels somewhat limited compared to more mature and flexible alternatives such as Copilot Studio or Azure AI Foundry.

This solution showcases the following concepts:

- Handle SSO authentication with Microsoft 365 Agents SDK.
- Integrate with AI Foundry Agent Service using the AI Foundry Javascript SDK.
- Register the Copilot Retrieval API as a tool
- Manage agent answers citations