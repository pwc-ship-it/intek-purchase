-- ============================================
-- 인텍플러스 구매관리 테이블 생성 스크립트
-- Supabase 대시보드 > SQL Editor 에서 실행하세요.
-- ============================================

-- 1. 품목 테이블 (여러 품목이 한 문서에 속하므로 doc_no는 중복 가능)
create table if not exists public.purchase_items (
    id           bigserial primary key,
    doc_no       text        not null,
    requester    text,
    p_name       text,
    p_code       text,
    name         text,
    spec         text,
    qty          integer default 0,
    uploaded_at  timestamptz default now()
);

-- 2. 조회 성능용 인덱스
create index if not exists idx_purchase_items_doc_no     on public.purchase_items (doc_no);
create index if not exists idx_purchase_items_uploaded   on public.purchase_items (uploaded_at desc);
create index if not exists idx_purchase_items_name       on public.purchase_items (name);

-- 3. Row Level Security (RLS) 활성화
alter table public.purchase_items enable row level security;

-- 4. 정책: 익명 사용자에게 읽기/쓰기 허용
--    (사내용이고 로그인 없이 쓸 거면 이 정책 사용)
--    나중에 인증 붙이면 정책을 authenticated 로 좁히세요.
drop policy if exists "anon_select" on public.purchase_items;
create policy "anon_select" on public.purchase_items
    for select using (true);

drop policy if exists "anon_insert" on public.purchase_items;
create policy "anon_insert" on public.purchase_items
    for insert with check (true);

drop policy if exists "anon_delete" on public.purchase_items;
create policy "anon_delete" on public.purchase_items
    for delete using (true);

drop policy if exists "anon_update" on public.purchase_items;
create policy "anon_update" on public.purchase_items
    for update using (true) with check (true);


-- ============================================
-- 5. (선택) DB 레벨 중복 방지 트리거
--    클라이언트 검사를 우회한 요청도 막고 싶다면 아래 트리거 사용.
--    이 트리거가 있으면 클라이언트가 한번이라도 이미 존재하는 doc_no로
--    insert 시도 시 PostgreSQL 에러코드 23505 (unique_violation) 발생.
-- ============================================
create or replace function public.check_doc_no_unique()
returns trigger
language plpgsql
as $$
begin
    -- 같은 insert 트랜잭션 안의 여러 행(같은 doc_no)은 허용
    -- 하지만 이미 테이블에 저장된 doc_no면 차단
    if exists (
        select 1 from public.purchase_items
        where doc_no = new.doc_no
          and id <> coalesce(new.id, -1)
    ) then
        raise exception 'duplicate doc_no: %', new.doc_no
            using errcode = '23505';
    end if;
    return new;
end;
$$;

drop trigger if exists trg_check_doc_no_unique on public.purchase_items;
create trigger trg_check_doc_no_unique
    before insert on public.purchase_items
    for each row
    execute function public.check_doc_no_unique();
