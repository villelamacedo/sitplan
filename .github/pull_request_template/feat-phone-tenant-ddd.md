# PR: feat(crm): extrair phone-normalizer, adicionar migration tenant_settings e script de migração dry‑run

Resumo

Extrai módulo de normalização de telefones (src/utils/phone-normalizer.ts), adiciona testes unitários (tests/phone-normalizer.test.ts), cria migration para tenant_settings + contacts.metadata e adiciona script de migração dry‑run (scripts/migrate_apply_ddd.js).

Problema resolvido

Remove hardcode de DDD no código, permite configuração de DDD por tenant e fornece ferramenta segura para aplicar migração nos contatos.

Arquivos incluídos

- src/utils/phone-normalizer.ts
- tests/phone-normalizer.test.ts
- scripts/migrate_apply_ddd.js
- db/migrations/2026-06-29-create-tenant-settings.sql

Como testar localmente

1. Rodar as migrations em um DB de staging (psql $DATABASE_URL -f db/migrations/2026-06-29-create-tenant-settings.sql).
2. Inserir seed em tenant_settings com DDD=11.
3. Rodar `npx jest` para garantir testes.
4. Executar `node scripts/migrate_apply_ddd.js --tenant <TENANT_UUID> --ddd 11 --dryrun` e revisar output.

Checklist

- [ ] Tests unitários adicionados e passando.
- [ ] Migration aplicada em staging sem erros.
- [ ] Script de migração validado via dry‑run.
- [ ] Documentação/README atualizada com API do normalizador e instruções de migração.
- [ ] Owner/ops notificado do plano de rollout.

Notas

- Recomenda-se rodar dry‑run e revisar amostra antes de aplicar a migração.
- Garantir backup/ snapshot da tabela `contacts` antes de qualquer operação em produção.
