# Multi-region failover test plan

Path under test: **Akamai ALB cloudlet → edge NLB → Envoy → (Avi GSLB DNS) → PrivateLink endpoint → service NLB → K8s service → pods**, with Region A as P0 and Region B as P1.

Diagram: [`multi-region-failover.drawio`](./multi-region-failover.drawio)

## Baseline configuration (15s end-to-end budget)

| Layer | Health check | Timeouts |
|---|---|---|
| Akamai ALB cloudlet | Liveness `/health/deep`, 10s interval, 5s timeout, fail on 5xx only | Connect 3s, read 15s, per-request retry on 5xx |
| Edge NLB | HTTP Envoy `/ready`, 5s × 2 | Dereg delay 30s, idle 350s, cross-zone on |
| Envoy (edge) | Active HC `/health/deep` 3s × 3 (healthy 3); outlier: 3 gateway errors, 30s ejection | Route 14s, per-try 4.5s, 2 retries, connect 1s, upstream idle 240s, max connection age 120s |
| Avi GSLB | HTTP monitor `/health/deep`, 5s × 3 | DNS TTL 15s, local → remote priority |
| PrivateLink endpoint | None | Idle drop 350s |
| Service NLB | HTTP `/healthz`, 5s × 2 | Dereg delay 30s, cross-zone on |
| Pods | Readiness 5s × 2 | App timeout 4s, preStop 20s, grace 45s, keepalive idle ≥ 300s |

## Prerequisites

- **Load generator** sending steady traffic through Akamai (external) and directly to Envoy (internal), with idempotent GET and non-idempotent POST mixes. Example: `k6` or `vegeta` at a fixed RPS. It must record per-request status, latency and the serving region (have the app return an `x-region` response header).
- **Envoy admin access** (`:9901/stats`, `/clusters`) or Envoy metrics in Prometheus/Grafana.
- **DNS check host** in the Envoy VPC: `watch -n1 dig +noall +answer app.gslb.internal`.
- **Dashboards**: NLB `HealthyHostCount` / `UnHealthyHostCount`, Avi GSLB service status, Akamai origin health and reports.
- **Fault tools**: `kubectl`, AWS FIS (AZ power interruption, network disruption), a security group you can edit, Avi controller access, Envoy fault filter or `tc netem` / Chaos Mesh for latency.
- **Run in a non-production environment first.** Announce each run, and keep a rollback command ready for every fault you inject.

## Envoy metrics to watch

| Stat | Meaning |
|---|---|
| `cluster.app.health_check.failure` / `.healthy` | Active health check results |
| `cluster.app.outlier_detection.ejections_active` | Hosts ejected by passive detection |
| `cluster.app.upstream_rq_retry` / `_success` | Retries attempted and recovered |
| `cluster.app.upstream_rq_per_try_timeout` | Per-try timeouts (4.5s) |
| `cluster.app.upstream_rq_timeout` | Route timeouts (14s) |
| `cluster.app.update_success` / `membership_change` | DNS refreshes and host-set changes |
| `cluster.app.upstream_cx_connect_timeout` | Connect timeouts (1s) |
| `cluster.app.upstream_cx_max_duration_reached` | Connections recycled at 120s |

## Test cases

Each case lists the fault, the expected behavior, and the pass criteria. "Errors" means non-2xx responses seen by the load generator.

### TC-01 Rolling deploy in Region A

- **Fault:** `kubectl rollout restart deploy/app -n <ns>` in Region A.
- **Expected:** Pods stop being ready, the preStop sleep lets the NLB deregister them, and in-flight requests finish within the grace period.
- **Pass:** 0 errors for GET. POST errors ≤ 0.01%. p99 latency below 2× baseline. No retries to P1.

### TC-02 Single pod crash

- **Fault:** `kubectl delete pod <pod> --grace-period=0 --force`.
- **Expected:** Connections to that pod reset. Envoy retries on `reset` to another host. The NLB removes the target within about 10s.
- **Pass:** GET errors = 0. Retries visible in `upstream_rq_retry_success`. No region failover.

### TC-03 Single AZ loss in Region A

- **Fault:** AWS FIS AZ power interruption, or a network disruption on one AZ's subnets.
- **Expected:** Envoy health checks fail for that AZ's PrivateLink ENI IP only (STRICT_DNS checks each IP). The other AZs keep serving. Cross-zone NLB keeps targets reachable.
- **Pass:** Traffic stays in Region A. Errors only during the first ~9s, and only on requests that could not be retried. One host ejected in `/clusters`.

### TC-04 All pods down in Region A

- **Fault:** `kubectl scale deploy/app --replicas=0` in Region A.
- **Expected:** The service NLB fails open, so a TCP-only check would pass. `/health/deep` fails, so Envoy marks P0 unhealthy in about 9s and moves traffic to P1. Avi flips the GSLB record after about 15s, plus up to 15s TTL.
- **Pass:** Envoy is fully on P1 in ≤ 12s. Retried GETs succeed through P1 (`previous_priorities`). Total errors limited to the first detection window. Akamai never times out at 15s.

### TC-05 Dependency failure in Region A (Redis or Kafka)

- **Fault:** Block pod egress to Redis or Kafka with a network policy, or stop the regional dependency.
- **Expected:** `/healthz` still passes, so pods stay in the NLB and are not restarted. `/health/deep` fails, so Envoy, Avi and Akamai fail the region over.
- **Pass:** No pod restarts (liveness must not be deep). Region failover within ~10s via Envoy. No cascading restarts in Region B.

### TC-06 PrivateLink endpoint unreachable

- **Fault:** Remove the inbound rule on the endpoint's security group, or detach its subnets.
- **Expected:** Connects fail or hang. The 1s connect timeout triggers retries to P1. Health checks eject P0 in about 9s.
- **Pass:** `upstream_cx_connect_timeout` rises. GET requests succeed via P1 within the 14s route timeout. No 15s Akamai timeouts.

### TC-07 Slow region (latency injection)

- **Fault:** Add 5s of latency to the Region A pods (Chaos Mesh NetworkChaos or the Envoy fault filter on a test listener).
- **Expected:** Each try hits the 4.5s per-try timeout and is retried, then moves to P1. Outlier detection ejects the slow hosts after 3 gateway errors.
- **Pass:** Requests finish in ≤ 14s, either through P1 or as a clean Envoy 504. Akamai never reports an origin read timeout.

### TC-08 Request budget enforcement

- **Fault:** A test endpoint that sleeps 20s.
- **Expected:** The app times out at 4s, Envoy's per-try timeout at 4.5s, and after 3 tries Envoy returns a 504 at about 14s.
- **Pass:** The client receives Envoy's 504 (check `server: envoy` / `x-envoy-*` headers), not an Akamai timeout error. Total time is between 13.5s and 14.5s.

### TC-09 Avi GSLB flip and TTL honored by Envoy

- **Fault:** Disable the Region A member in the Avi GSLB service.
- **Expected:** `dig` shows the Region B answer with TTL ≤ 15. Envoy re-resolves within one TTL (`respect_dns_ttl`), P0 now points at Region B, and `membership_change` increments.
- **Pass:** The Envoy host set changes ≤ 20s after the flip. Long-lived connections move within 120s (`upstream_cx_max_duration_reached`). No errors.

### TC-10 DNS resolution failure

- **Fault:** Block Envoy's access to the resolver (VPC DNS) with a network policy, or point Envoy at a dead resolver on a test instance.
- **Expected:** Envoy keeps its last resolved hosts and retries DNS on the `dns_failure_refresh_rate` backoff (2–10s).
- **Pass:** No traffic loss while DNS is down. Hosts are not removed. Recovery is automatic once DNS returns.

### TC-11 Envoy edge pod drain

- **Fault:** `kubectl delete pod` on one Envoy edge pod (preStop calls `/healthcheck/fail`, then sleeps 25s).
- **Expected:** The edge NLB marks the pod unhealthy in about 10s and deregisters it. Akamai connections drain without resets.
- **Pass:** 0 external errors. Akamai origin errors flat.

### TC-12 Full edge loss in Region A (Akamai failover)

- **Fault:** Scale the Envoy edge deployment to 0 in Region A, or block the edge NLB's listener.
- **Expected:** Akamai's per-request retry on 5xx or connect failure sends traffic to Region B at once. The liveness test marks Region A down after about 10–20s.
- **Pass:** External error rate returns to baseline within ≤ 30s. Afterwards, no requests to Region A until it is healthy again.

### TC-13 Both regions unhealthy (Akamai signal)

- **Fault:** Fail `/health/deep` in both regions on a test environment.
- **Expected:** The Envoy `health_check` filter drops below 50% healthy hosts and returns 503 on the health route. Akamai marks the origin down and serves its fail action (error page or failover origin).
- **Pass:** Akamai shows the configured fail action. Envoy does not report itself healthy while it has no healthy backend.

### TC-14 Failback without flapping

- **Fault:** Restore Region A after TC-04 or TC-05.
- **Expected:** Envoy needs 3 healthy checks (about 9s) before P0 takes traffic again. Avi and Akamai fail back after their own healthy thresholds.
- **Pass:** One clean switch back with no oscillation. Toggle the fault every 5s for 1 minute: Envoy does not switch regions more than once per cycle.

### TC-15 Idle connection survival

- **Fault:** None. Open a pooled connection and keep it idle for 330s, then send a request.
- **Expected:** Envoy closes the upstream connection at 240s, before PrivateLink's 350s drop, and the request uses a fresh connection.
- **Pass:** No resets or 503s on the first request after the idle period. `upstream_cx_idle_timeout` increments.

### TC-16 Long-lived connection rebalancing

- **Fault:** Run gRPC or HTTP/2 streams, then scale Region A pods from 3 to 6.
- **Expected:** Connections recycle at 120s, so the new pods receive traffic.
- **Pass:** Within 3 minutes, request share per pod is within ±20% of even.

## Results log

| Case | Date | Env | Detection time | Errors (count / %) | Max latency | Pass? | Notes |
|---|---|---|---|---|---|---|---|
| TC-01 | | | | | | | |
| TC-02 | | | | | | | |
| TC-03 | | | | | | | |
| TC-04 | | | | | | | |
| TC-05 | | | | | | | |
| TC-06 | | | | | | | |
| TC-07 | | | | | | | |
| TC-08 | | | | | | | |
| TC-09 | | | | | | | |
| TC-10 | | | | | | | |
| TC-11 | | | | | | | |
| TC-12 | | | | | | | |
| TC-13 | | | | | | | |
| TC-14 | | | | | | | |
| TC-15 | | | | | | | |
| TC-16 | | | | | | | |
