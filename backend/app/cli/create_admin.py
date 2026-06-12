import argparse
import asyncio

from sqlalchemy import select

from app.core.database import async_session_maker
from app.core.security import get_password_hash
from app.models.user import User


async def create_admin(username: str, password: str, name: str, email: str | None) -> None:
    async with async_session_maker() as session:
        existing = await session.execute(select(User).where(User.username == username))
        if existing.scalar_one_or_none():
            raise SystemExit(f"User already exists: {username}")

        session.add(
            User(
                username=username,
                name=name,
                email=email,
                hashed_password=get_password_hash(password),
                role="admin",
                is_active=True,
            )
        )
        await session.commit()


def main() -> None:
    parser = argparse.ArgumentParser(description="Create an initial admin user.")
    parser.add_argument("--username", required=True)
    parser.add_argument("--password", required=True)
    parser.add_argument("--name", required=True)
    parser.add_argument("--email")
    args = parser.parse_args()

    asyncio.run(create_admin(args.username, args.password, args.name, args.email))


if __name__ == "__main__":
    main()

