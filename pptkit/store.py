# -*- coding: utf-8 -*-
"""产物仓库：登记 / 校验 / 列表 / 取用。

每个产物形如 {"id", "name", "kind", "created_at", "bytes", "path", "sha256", "spec"}
kind ∈ {pptx, pdf, png, json}
"""
import hashlib
import json
import os
import threading
import time

_lock = threading.Lock()


def _sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


class Store:
    def __init__(self, root):
        self.root = os.path.abspath(root)
        os.makedirs(self.root, exist_ok=True)
        self._index_path = os.path.join(self.root, "index.json")
        self._items = []
        self._load()

    def _load(self):
        if os.path.exists(self._index_path):
            try:
                with open(self._index_path, encoding="utf-8") as f:
                    self._items = json.load(f)
            except Exception:
                self._items = []
        # 丢掉已被删除的文件
        self._items = [it for it in self._items if os.path.exists(it["path"])]

    def _flush(self):
        tmp = self._index_path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(self._items, f, ensure_ascii=False, indent=2)
        os.replace(tmp, self._index_path)

    def new_id(self, prefix="deck"):
        import secrets
        return "%s_%s_%s" % (prefix, time.strftime("%Y%m%d_%H%M%S"), secrets.token_hex(8))

    def register(self, name, kind, path, spec=None, item_id=None, validate=True, compute_sha=False):
        path = os.path.abspath(path)
        if validate and not os.path.exists(path):
            raise FileNotFoundError(path)
        rec = {
            "id": item_id or self.new_id(kind),
            "name": name,
            "kind": kind,
            "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
            "bytes": os.path.getsize(path) if os.path.exists(path) else 0,
            "path": path,
            "sha256": _sha256(path) if (compute_sha and os.path.exists(path)) else "",
            "spec": spec,
        }
        with _lock:
            self._items = [it for it in self._items if it["id"] != rec["id"]]
            self._items.insert(0, rec)
            self._flush()
        return rec

    def register_many(self, records, validate=True):
        """批量登记产物，加锁一次并只持久化写入一次 index.json。"""
        recs = []
        for r in records:
            path = os.path.abspath(r["path"])
            if validate and not os.path.exists(path):
                continue
            rec = {
                "id": r.get("item_id") or self.new_id(r["kind"]),
                "name": r["name"],
                "kind": r["kind"],
                "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "bytes": os.path.getsize(path) if os.path.exists(path) else 0,
                "path": path,
                "sha256": "",
                "spec": r.get("spec"),
            }
            recs.append(rec)
        if not recs:
            return []
        with _lock:
            new_ids = {rc["id"] for rc in recs}
            self._items = [it for it in self._items if it["id"] not in new_ids]
            for rc in reversed(recs):
                self._items.insert(0, rc)
            self._flush()
        return recs

    def list(self, kind=None, limit=50):
        with _lock:
            items = list(self._items)
        if kind:
            items = [it for it in items if it["kind"] == kind]
        return items[:limit]

    def get(self, item_id):
        key = str(item_id or "").strip()
        key_no_ext = key.removesuffix(".pptx").removesuffix(".pdf").removesuffix(".png")
        with _lock:
            for it in self._items:
                if it["id"] in (key, key_no_ext):
                    return it
                if it["name"] in (key, key_no_ext, key + ".pptx"):
                    return it
        return None

    def update_spec(self, item_id, spec):
        """更新某个产物对应的 spec。"""
        rec = self.get(item_id)
        if not rec:
            return None
        with _lock:
            rec["spec"] = spec
            self._flush()
        return rec

    def delete(self, item_id):
        rec = self.get(item_id)
        if not rec:
            return False
        with _lock:
            self._items = [it for it in self._items if it["id"] != rec["id"]]
            self._flush()
        try:
            if os.path.exists(rec["path"]):
                os.remove(rec["path"])
            # 清理关联的预览截图目录
            base_slug = "".join(c for c in os.path.splitext(rec["name"])[0] if c.isalnum() or c in "-_ ").strip().replace(" ", "_")
            pv_dir = os.path.join(os.path.dirname(rec["path"]), "preview", base_slug)
            if os.path.exists(pv_dir):
                import shutil
                shutil.rmtree(pv_dir, ignore_errors=True)
        except OSError:
            pass
        return True


    def path_of(self, item_id):
        rec = self.get(item_id)
        if not rec:
            # 尝试直接在仓库根目录或直接路径寻找
            cand = os.path.join(self.root, item_id)
            if not os.path.exists(cand):
                cand = os.path.join(self.root, item_id + ".pptx")
            if os.path.exists(cand):
                return cand
            raise KeyError("产物不存在: %s" % item_id)
        if not os.path.exists(rec["path"]):
            raise FileNotFoundError(rec["path"])
        return rec["path"]
