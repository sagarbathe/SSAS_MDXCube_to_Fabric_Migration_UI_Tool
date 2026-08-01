"""
One-off helper: binds a semantic model's SQL data source to a Fabric
"cloud connection" object via the REST API, for cases where the connection
doesn't show up in the semantic model's Settings > Gateway connections
picker in the portal (a known Fabric UI caching/visibility issue).

Usage:
    .venv\\Scripts\\python.exe scripts\\bind_connection.py ^
        --env-file config\\.env ^
        --semantic-model-name AutoInsuranceCubeDemo ^
        --lakehouse-name AutoInsuranceLakehouse

Run with --list-only first if you just want to see every connection the
service principal can see (to confirm your new connection actually exists
and is visible to it) without binding anything.
"""
from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))

from ssas_fabric_migrator.cli.orchestrator import load_env
from ssas_fabric_migrator.deploy.fabric_client import FabricClient


def main():
    parser = argparse.ArgumentParser(description="Bind a semantic model's SQL data source to a Fabric connection")
    parser.add_argument("--env-file", default="config/.env")
    parser.add_argument("--semantic-model-name")
    parser.add_argument("--lakehouse-name")
    parser.add_argument("--connection-id", help="Skip auto-matching and use this connection object ID directly")
    parser.add_argument("--list-only", action="store_true", help="Just list visible connections and exit")
    args = parser.parse_args()

    env = load_env(args.env_file)
    client = FabricClient(env["FABRIC_TENANT_ID"], env["FABRIC_CLIENT_ID"], env["FABRIC_CLIENT_SECRET"])
    workspace_id = env["FABRIC_WORKSPACE_ID"]

    connections = client.list_connections()
    print(f"Service principal can see {len(connections)} connection(s):")
    for c in connections:
        details = c.get("connectionDetails", {})
        print(f"  - id={c['id']}  displayName={c.get('displayName')!r}  type={details.get('type')}  path={details.get('path')!r}  connectivityType={c.get('connectivityType')}")

    if args.list_only:
        return

    if not args.semantic_model_name or not args.lakehouse_name:
        print("\n--semantic-model-name and --lakehouse-name are required unless --list-only is given.")
        sys.exit(1)

    lakehouse = client.find_item(workspace_id, args.lakehouse_name, "Lakehouse")
    if not lakehouse:
        print(f"Lakehouse '{args.lakehouse_name}' not found in workspace.")
        sys.exit(1)
    endpoint = client.get_lakehouse_sql_endpoint(workspace_id, lakehouse["id"])
    if not endpoint:
        print("Could not resolve the Lakehouse's SQL analytics endpoint (it may still be provisioning).")
        sys.exit(1)
    sql_path = f"{endpoint};{args.lakehouse_name}"
    print(f"\nTarget SQL data source path: {sql_path!r}")

    model = client.find_item(workspace_id, args.semantic_model_name, "SemanticModel")
    if not model:
        print(f"Semantic model '{args.semantic_model_name}' not found in workspace.")
        sys.exit(1)

    if args.connection_id:
        connection_id = args.connection_id
    else:
        matches = [c for c in connections if c.get("connectionDetails", {}).get("type") == "SQL"
                   and endpoint.split(",")[0].split(":")[0].lower() in (c.get("connectionDetails", {}).get("path") or "").lower()]
        if not matches:
            print("\nNo visible connection's path matches this Lakehouse's SQL endpoint host.")
            print("Re-run with --connection-id <id> using one of the IDs printed above (pick the SQL connection you created),")
            print("or double-check the connection was created with the exact 'SQL connection string' from the Lakehouse's SQL analytics endpoint page.")
            sys.exit(1)
        if len(matches) > 1:
            print("\nMultiple candidate connections matched - pass --connection-id explicitly:")
            for c in matches:
                print(f"  - {c['id']}  {c.get('displayName')!r}")
            sys.exit(1)
        connection_id = matches[0]["id"]
        print(f"Matched connection: {matches[0].get('displayName')!r} ({connection_id})")

    print(f"\nBinding semantic model '{args.semantic_model_name}' ({model['id']}) to connection {connection_id} ...")
    client.bind_semantic_model_connection(workspace_id, model["id"], connection_id, sql_path)
    print("Bind succeeded. Now trigger a refresh (Fabric portal 'Refresh now', or re-run the UI tool's 'deploy-model' step).")


if __name__ == "__main__":
    main()
