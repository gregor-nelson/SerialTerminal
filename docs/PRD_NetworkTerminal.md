# Network Terminal - Product Requirements Document

## TCP/UDP Terminal Application

**Document Version:** 2.0
**Date:** 2026-01-13
**Status:** Draft
**Parent Project:** SerialTerminal

---

## Quick Navigation

### Implementation Phases

| Phase | Document | Owner | Status |
|-------|----------|-------|--------|
| **Phase 0** | [Project Setup](./phases/phase-0-project-setup.md) | Any Dev | Not Started |
| **Stream A** | [Core Network Layer](./phases/stream-a-core-network.md) | Dev 1 | Not Started |
| **Stream B** | [UI Layer](./phases/stream-b-ui-layer.md) | Dev 2 | Not Started |
| **Stream C** | [Config & Data Layer](./phases/stream-c-config-data.md) | Dev 3 | Not Started |
| **Phase 4** | [Integration](./phases/phase-4-integration.md) | All Devs | Not Started |
| **Phase 5** | [Testing & Polish](./phases/phase-5-testing.md) | All Devs | Not Started |

### Reference Documents

| Document | Description |
|----------|-------------|
| [Architecture Analysis](./reference/architecture-analysis.md) | Current Serial Terminal codebase breakdown |
| [Reusability Matrix](./reference/reusability-matrix.md) | What to copy, adapt, or replace |
| [Interface Definitions](./reference/interface-definitions.md) | API contracts between components |
| [File Structure](./reference/file-structure.md) | Target project organization |
| [Testing Requirements](./reference/testing-requirements.md) | Unit, integration, manual tests |

---

## Executive Summary

### Objective

Create a standalone **Network Terminal Application** for TCP and UDP connections that mirrors the design, layout, and feature set of the existing Serial Terminal application. The goal is to maximize code reuse while creating an independent, maintainable application.

### Key Goals

| Goal | Target |
|------|--------|
| Code reuse from Serial Terminal | **85%+** |
| UI/UX consistency | Identical to Serial Terminal |
| Protocol support | TCP Client, TCP Server, UDP |
| Independence | Standalone application |
| Parallel development | 3 concurrent work streams |

### Scope

| In Scope | Out of Scope |
|----------|--------------|
| TCP Client connections | Serial port support |
| TCP Server (listen mode) | Virtual port management |
| UDP Sender/Receiver | Baud rate detection |
| UDP Multicast support | Hardware flow control |
| Connection history/favorites | Moxa device integration |
| Split-pane terminal display | Windows registry scanning |
| NMEA message color-coding | |

---

## Development Workflow

### Phase Dependencies

```
┌─────────────────────────────────────────────────────────────────┐
│                    PHASE 0: Project Setup                       │
│                    [phase-0-project-setup.md]                   │
│                    Single developer, prerequisite               │
└─────────────────────────────────────────────────────────────────┘
                              │
        ┌─────────────────────┼─────────────────────┐
        │                     │                     │
        ▼                     ▼                     ▼
┌───────────────┐    ┌───────────────┐    ┌───────────────┐
│   STREAM A    │    │   STREAM B    │    │   STREAM C    │
│  Core Network │    │   UI Layer    │    │  Config/Data  │
│   [Dev 1]     │    │   [Dev 2]     │    │   [Dev 3]     │
│               │    │               │    │               │
│ Can run in    │    │ Can run in    │    │ Can run in    │
│ PARALLEL      │    │ PARALLEL      │    │ PARALLEL      │
└───────────────┘    └───────────────┘    └───────────────┘
        │                     │                     │
        └─────────────────────┼─────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    PHASE 4: Integration                         │
│                    [phase-4-integration.md]                     │
│                    All developers collaborate                   │
└─────────────────────────────────────────────────────────────────┘
                              │
                              ▼
┌─────────────────────────────────────────────────────────────────┐
│                    PHASE 5: Testing & Polish                    │
│                    [phase-5-testing.md]                         │
│                    All developers collaborate                   │
└─────────────────────────────────────────────────────────────────┘
```

### Estimated Effort

| Phase | Tasks | New Lines | Adapted Lines | Copied Lines |
|-------|-------|-----------|---------------|--------------|
| Phase 0 | 4 | 20 | 0 | 850 |
| Stream A | 5 | 400 | 0 | 0 |
| Stream B | 4 | 200 | 400 | 1700 |
| Stream C | 3 | 250 | 50 | 150 |
| Phase 4 | 3 | 50 | 100 | 0 |
| Phase 5 | 4 | 50 | 50 | 0 |
| **Total** | **23** | **970** | **600** | **2700** |

---

## Getting Started

### For New Developers

1. **Read** [Architecture Analysis](./reference/architecture-analysis.md) to understand the existing codebase
2. **Check** [Reusability Matrix](./reference/reusability-matrix.md) for what can be copied
3. **Review** [Interface Definitions](./reference/interface-definitions.md) for API contracts
4. **Start** with [Phase 0: Project Setup](./phases/phase-0-project-setup.md)

### For Parallel Development

After Phase 0 is complete, developers can pick up any stream:

- **Developer 1**: Start [Stream A: Core Network](./phases/stream-a-core-network.md)
- **Developer 2**: Start [Stream B: UI Layer](./phases/stream-b-ui-layer.md)
- **Developer 3**: Start [Stream C: Config/Data](./phases/stream-c-config-data.md)

### Task Tracking

Each phase document contains:
- Detailed task breakdown with acceptance criteria
- Status checkboxes for tracking progress
- Links to related documents
- Code examples and snippets

---

## Document Index

### Phase Documents

| Document | Purpose | Prerequisites |
|----------|---------|---------------|
| [phase-0-project-setup.md](./phases/phase-0-project-setup.md) | Create project skeleton, copy reusable files | None |
| [stream-a-core-network.md](./phases/stream-a-core-network.md) | Implement NetworkWorker classes | Phase 0 |
| [stream-b-ui-layer.md](./phases/stream-b-ui-layer.md) | Adapt UI components | Phase 0 |
| [stream-c-config-data.md](./phases/stream-c-config-data.md) | Create config and data management | Phase 0 |
| [phase-4-integration.md](./phases/phase-4-integration.md) | Wire up all components | Streams A, B, C |
| [phase-5-testing.md](./phases/phase-5-testing.md) | Test and polish | Phase 4 |

### Reference Documents

| Document | Purpose |
|----------|---------|
| [architecture-analysis.md](./reference/architecture-analysis.md) | Detailed breakdown of Serial Terminal |
| [reusability-matrix.md](./reference/reusability-matrix.md) | Copy/adapt/replace decisions |
| [interface-definitions.md](./reference/interface-definitions.md) | API contracts |
| [file-structure.md](./reference/file-structure.md) | Target project layout |
| [testing-requirements.md](./reference/testing-requirements.md) | All test cases |

---

## Version History

| Version | Date | Author | Changes |
|---------|------|--------|---------|
| 1.0 | 2026-01-13 | Claude | Initial monolithic PRD |
| 2.0 | 2026-01-13 | Claude | Split into linked documents |

---

## Related Links

- **Source Repository**: SerialTerminal (parent project)
- **Target Repository**: NetworkTerminal (new project)

---

[Back to Top](#network-terminal---product-requirements-document)
