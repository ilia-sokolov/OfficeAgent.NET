# Hosted gateway reference boundary

This sample hosts OfficeAgent's MCP endpoint behind ASP.NET Core authentication and an
`IConnectionAccessPolicy`. Alice can use only `connection-a`; Bob can use only
`connection-b`. The policy runs before provider lookup, registration, document reads, edits,
creation, or registration removal. `list_connections` returns only the caller's connection.
The same trusted ASP.NET Core identity is copied into apply and merge receipt `actor` fields;
revision authors remain separate display metadata supplied by the plan.

Run it with:

```powershell
dotnet run --project samples/HostedGateway
```

The sample accepts `Authorization: Bearer alice-token` and `Authorization: Bearer bob-token` so the
boundary can be exercised without an identity tenant. This authentication handler is only a
demonstration. Replace it with validated JWT bearer authentication and map stable issuer and
subject claims in the policy before deployment. Do not accept a caller name from MCP tool
arguments or another model-controlled field.

`HostedGateway__ConnectionARoot` and `HostedGateway__ConnectionBRoot` can override the two
filesystem roots. The defaults are private folders under the sample's output directory.

ASP.NET request/response body logging is not enabled. The sample raises ASP.NET Core and MCP
framework logs to `Warning`, so bearer tokens, plans, and document content do not enter
application logs. Keep that property when adding observability: record authenticated subject,
operation, connection, outcome, and receipt hashes, but never authorization headers, source
paths supplied by callers, plan bodies, or document bytes.

This is a reference authorization boundary, not a hosted product. A reusable hosted tier also
needs tenant provisioning, quotas, encryption, secret rotation, retention controls, regional
deployment, billing, abuse controls, and operational support.

Keep hosted-product promotion behind a deployment evidence gate. Record at least two independent
deployments with different identity configurations, each repeating the two-principal isolation
test against its real connection policy and documenting incident response, secret rotation, data
retention, and quota behavior. Treat a third deployment as required when the first two use the
same cloud or identity provider.
