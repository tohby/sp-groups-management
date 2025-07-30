import * as React from "react";
import styles from "./GroupManagement.module.scss";
import { sp } from "@pnp/sp/presets/all";
import Member from "./member";
import {
  CompactPeoplePicker,
  DefaultButton,
  IPersonaProps,
  MessageBar,
  MessageBarType,
  Overlay,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack,
} from "@fluentui/react";

interface IGroupDetailProps {
  groupId: number;
}

const GroupDetail: React.FC<IGroupDetailProps> = ({ groupId }) => {
  const { useEffect, useState } = React;
  const [groupDetail, setGroupDetail] = useState<any>([]);
  const [groupMembers, setGroupMembers] = useState<any>([]);
  const [isLoading, setIsLoading] = React.useState<boolean>(false);
  const [selectedUsers, setSelectedUsers] = useState<IPersonaProps[]>([]);
  const [message, setMessage] = useState<{
    type: "success" | "error";
    text: string;
  } | null>(null);
  const fetchGroupDetail = async () => {
    const [response, members] = await Promise.all([
      sp.web.siteGroups.getById(groupId).get(),
      sp.web.siteGroups.getById(groupId).users(),
    ]);
    setGroupMembers(members);
    setGroupDetail(response);
  };

  useEffect(() => {
    if (groupId) {
      fetchGroupDetail().catch((error) => {
        console.error("Error fetching group details:", error);
      });
    }
  }, [groupId]);

  const removeUserFromGroup = async (userId: number) => {
    setIsLoading(true);
    try {
      await sp.web.siteGroups.getById(groupId).users.removeById(userId);
      fetchGroupDetail().catch((error) => {
        console.error("Error fetching group details:", error);
      });
    } catch (error) {
      console.error("Error removing user from group:", error);
    } finally {
      setIsLoading(false);
    }
  };

  const onFilterChanged = async (
    filterText: string,
    currentPersonas: IPersonaProps[],
    limitResults?: number
  ): Promise<IPersonaProps[]> => {
    if (!filterText) return [];

    try {
      const results = await sp.web.siteUsers
        .filter(
          `substringof('${filterText}', Title) or substringof('${filterText}', Email)`
        )
        .top(10)();

      return results.map((user) => ({
        text: user.Title,
        secondaryText: user.UserPrincipalName || user.Email,
        id: user.Id.toString(),
        key: user.LoginName,
      }));
    } catch (error) {
      console.error("Error fetching user suggestions:", error);
      return [];
    }
  };

  const addUsersToGroup = async () => {
    setIsLoading(true);
    try {
      for (const user of selectedUsers) {
        console.log(user);

        const ensuredUser = await sp.web.ensureUser(user.secondaryText!);
        await sp.web.siteGroups
          .getById(groupId)
          .users.add(ensuredUser.data.LoginName);
      }
      fetchGroupDetail().catch((error) => {
        console.error("Error fetching group details:", error);
      });
      setSelectedUsers([]);
    } catch (error) {
      console.error("Error adding users to group:", error);
    } finally {
      setIsLoading(false);
    }
  };

  return (
    <div className={styles.groupDetail}>
      {groupDetail.Title} (Id: {groupId})
      <p>
        Group owner: <b>{groupDetail.OwnerTitle}</b>
      </p>
      <Stack horizontal styles={{ root: { width: "100%" } }}>
        <Stack.Item grow>
          <CompactPeoplePicker
            onResolveSuggestions={onFilterChanged}
            selectedItems={selectedUsers}
            onChange={(items: any) => setSelectedUsers(items)}
          />
        </Stack.Item>
        <PrimaryButton
          text="Add"
          iconProps={{ iconName: "AddFriend" }}
          onClick={addUsersToGroup}
        />
      </Stack>
      {message && (
        <MessageBar
          messageBarType={
            message.type === "success"
              ? MessageBarType.success
              : MessageBarType.error
          }
          onDismiss={() => setMessage(null)}
        >
          {message.text}
        </MessageBar>
      )}
      {groupMembers.length > 0 ? (
        <div className={styles.membersList}>
          {groupMembers.map((member: any) => (
            <Stack
              key={member.Id}
              horizontal
              className={styles.memberItem}
              horizontalAlign="space-between"
              verticalAlign="center"
            >
              <Member name={member.Title} email={member.UserPrincipalName} />
              <DefaultButton
                text="remove"
                iconProps={{ iconName: "UserRemove" }}
                onClick={() => removeUserFromGroup(member.Id)}
              />
            </Stack>
          ))}
        </div>
      ) : (
        <p>No members found.</p>
      )}
      {isLoading && (
        <Overlay>
          <Spinner
            label="Please hold on; this may take a moment..."
            size={SpinnerSize.large}
            styles={{
              root: {
                position: "absolute",
                top: "50%",
                left: "50%",
                transform: "translate(-50%, -50%)",
              },
            }}
          />
        </Overlay>
      )}
    </div>
  );
};

export default GroupDetail;
