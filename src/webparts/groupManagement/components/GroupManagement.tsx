import { Stack, Text, Nav, INavLink } from "@fluentui/react";
import { ISiteGroupInfo, sp } from "@pnp/sp/presets/all";
import * as React from "react";
import GroupDetail from "./GroupDetail";

const GroupManagement = () => {
  const { useEffect, useState } = React;
  const [groups, setGroups] = useState<ISiteGroupInfo[]>([]);
  const [selectedGroupKey, setSelectedGroupKey] = useState<string>("0");
  const fetchGroups = async () => {
    const response = await sp.web.siteGroups.get();
    setGroups(response);
    setSelectedGroupKey(response.length > 0 ? response[0].Id.toString() : "1");
  };
  useEffect(() => {
    fetchGroups().catch((error) => {
      console.error("Error fetching groups:", error);
    });
  }, []);

  const onGroupSelected = (
    ev: React.MouseEvent<HTMLElement>,
    item?: INavLink
  ) => {
    if (item && item.key) {
      setSelectedGroupKey(item.key);
    }
  };
  return (
    <div>
      <Stack horizontal styles={{ root: { height: "100vh" } }}>
        <Stack
          styles={{
            root: { width: "20%" },
          }}
        >
          <Text variant="xLarge" block>
            Groups
          </Text>
          <div style={{ marginTop: "20px" }}>
            <Nav
              selectedKey={selectedGroupKey}
              onLinkClick={onGroupSelected}
              groups={[
                {
                  links: groups.map((group) => ({
                    name: group.Title,
                    key: group.Id.toString(), 
                    url: "#", 
                  })),
                },
              ]}
            />
          </div>
        </Stack>
        <Stack grow styles={{ root: { padding: 20 } }}>
          <GroupDetail groupId={Number(selectedGroupKey)} />
        </Stack>
      </Stack>
    </div>
  );
};

export default GroupManagement;
